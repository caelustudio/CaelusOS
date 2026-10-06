#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPTOS VBA 工具链 —— 把 .pptm 里的 VBA 源码抽出来做版本管理，并支持重建回 .pptm。

用法：
    python3 pptos_vba.py extract <in.pptm> <src_dir>      # 抽取 VBA 源码（UTF-8）
    python3 pptos_vba.py build   <src_dir> <in.pptm> <out.pptm> [--keep-srp]
    python3 pptos_vba.py verify  <pptm> <src_dir>         # 比对 pptm 内宏与源码是否一致
    python3 pptos_vba.py lint    <src_dir>                # 粗检 VBA 语法块平衡/引号
    python3 pptos_vba.py keys                             # 生成混淆后的 app_key 密文

实现要点（MS-OVBA / MS-CFB）：
  * 模块流（VBA/<模块名>）= CompressedContainer(源码) + PerformanceCache(已编译 p-code)。
    重建时只写「源码」部分、丢弃 p-code，宿主必须重新编译 → 改动真正生效。
  * 源文件以 UTF-8 存仓库，写回时转成项目代码页（PROJECTCODEPAGE，本工程为 936/GBK）。
  * 压缩器采用「全字面量」编码（合法但压缩率低），因此新流体积必然小于原流，
    可以原地覆写扇区链，不必重排整个 OLE。
"""
import argparse
import os
import re
import struct
import sys
import zipfile

# --------------------------------------------------------------------------
# MS-OVBA 压缩容器
# --------------------------------------------------------------------------
CHUNK_MAX_RAW = 3600  # 全字面量编码下不超 12bit 上限的安全值
SIGNATURE = 0x01


def _bit_count(pos_in_chunk: int) -> int:
    """MS-OVBA 2.4.1.3.19.1：随块内位置变化的 offset/length 位分配。"""
    if pos_in_chunk <= 0:
        return 4
    return max((pos_in_chunk - 1).bit_length(), 4)


def ovba_compress(data: bytes) -> bytes:
    """把字节串编码成 MS-OVBA CompressedContainer（贪心 LZ77 压缩）。

    块头 12 位字段 = 数据长度 - 1（块长 = 字段 + 3，其中含 2 字节头）。
    每个解压块最多 4096 字节；回引不得跨块。
    """
    out = bytearray([SIGNATURE])
    if not data:
        out += struct.pack('<H', 0x0000 | (0b011 << 12) | (1 << 15)) + b'\x00'
        return bytes(out)

    for coff in range(0, len(data), 4096):
        chunk = data[coff:coff + 4096]
        n = len(chunk)
        body = bytearray()
        heads = {}                     # 3 字节前缀 -> 最近出现的位置
        i = 0
        while i < n:
            flag_index = len(body)
            body.append(0x00)
            flag = 0
            for bit in range(8):
                if i >= n:
                    break
                bc = _bit_count(i)
                length_mask = 0xFFFF >> bc
                offset_mask = (~length_mask) & 0xFFFF
                max_len = min(length_mask + 3, n - i)
                max_off = (offset_mask >> (16 - bc)) + 1
                best_len, best_off = 0, 0
                if i > 0 and max_len >= 3 and i + 2 < n:
                    lo = max(0, i - max_off)
                    for start in reversed(heads.get(chunk[i:i + 3], [])):
                        if start < lo:
                            break
                        L = 0
                        while L < max_len and chunk[start + L] == chunk[i + L]:
                            L += 1
                        if L > best_len:
                            best_len, best_off = L, i - start
                            if L == max_len:
                                break
                if best_len >= 3:
                    token = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                    body += struct.pack('<H', token & 0xFFFF)
                    flag |= (1 << bit)
                    for k in range(i, i + best_len):
                        if k + 2 < n:
                            h = chunk[k:k + 3]
                            lst = heads.setdefault(h, [])
                            lst.append(k)
                            if len(lst) > 24:
                                del lst[0]
                    i += best_len
                else:
                    if i + 2 < n:
                        h = chunk[i:i + 3]
                        lst = heads.setdefault(h, [])
                        lst.append(i)
                        if len(lst) > 24:
                            del lst[0]
                    body.append(chunk[i])
                    i += 1
            body[flag_index] = flag
        header = ((len(body) - 1) & 0x0FFF) | (0b011 << 12) | (1 << 15)
        out += struct.pack('<H', header) + body
    return bytes(out)


def ovba_compress_literal(data: bytes) -> bytes:
    """全字面量编码（压缩率差，用于需要把流撑到 4096 字节以上的场景）。

    每块数据不超过 3600 字节：块头 12 位字段 ≤ 4095 的硬限制。
    """
    out = bytearray([SIGNATURE])
    if not data:
        out += struct.pack('<H', 0x0000 | (0b011 << 12) | (1 << 15)) + b'\x00'
        return bytes(out)
    step = 3600
    for off in range(0, len(data), step):
        piece = data[off:off + step]
        body = bytearray()
        for i in range(0, len(piece), 8):
            body.append(0x00)
            body.extend(piece[i:i + 8])
        header = ((len(body) - 1) & 0x0FFF) | (0b011 << 12) | (1 << 15)
        out += struct.pack('<H', header) + body
    return bytes(out)


def ovba_decompress(container: bytes) -> bytes:
    """MS-OVBA 解压（自实现，用于构建后自检）。"""
    if not container or container[0] != SIGNATURE:
        raise ValueError('invalid signature byte %02X' % (container[0] if container else -1))
    pos = 1
    out = bytearray()
    n = len(container)
    while pos + 2 <= n:
        header = struct.unpack_from('<H', container, pos)[0]
        chunk_size = (header & 0x0FFF) + 3
        signature = (header >> 12) & 0x07
        flag = (header >> 15) & 0x01
        if signature != 0b011:
            raise ValueError('invalid chunk signature at %d' % pos)
        pos += 2
        end = min(n, pos + chunk_size - 2)
        if flag == 0:                        # 未压缩块：4096 字节原样
            out += container[pos:end]
            pos = end
        else:                                # 压缩块：逐 token 还原
            chunk_start = len(out)
            while pos < end:
                flags = container[pos]
                pos += 1
                for bit in range(8):
                    if pos >= end:
                        break
                    if flags & (1 << bit):
                        if pos + 2 > end:
                            break
                        token = struct.unpack_from('<H', container, pos)[0]
                        pos += 2
                        # CopyToken 的位分配随块内位置变化（MS-OVBA 2.4.1.3.19.1）
                        diff = max(1, len(out) - chunk_start)
                        bit_count = max((diff - 1).bit_length(), 4)
                        length_mask = 0xFFFF >> bit_count
                        offset_mask = (~length_mask) & 0xFFFF
                        length = (token & length_mask) + 3
                        offset = ((token & offset_mask) >> (16 - bit_count)) + 1
                        start = len(out) - offset
                        if start < 0:
                            raise ValueError('bad copy token offset')
                        for k in range(length):
                            out.append(out[start + k])
                    else:
                        out.append(container[pos])
                        pos += 1
            pos = end
    return bytes(out)


# --------------------------------------------------------------------------
# MS-CFB（OLE 复合文档）最小实现：只做「原地替换流内容」
# --------------------------------------------------------------------------
class Cfb:
    def __init__(self, data: bytes):
        self.raw = bytearray(data)
        if self.raw[:8] != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            raise ValueError('not a CFB file')
        self.sector_shift = struct.unpack_from('<H', self.raw, 0x1E)[0]
        self.mini_shift = struct.unpack_from('<H', self.raw, 0x20)[0]
        self.sector_size = 1 << self.sector_shift
        self.mini_cutoff = struct.unpack_from('<L', self.raw, 0x38)[0]
        self.dir_start = struct.unpack_from('<L', self.raw, 0x30)[0]
        self.n_fat = struct.unpack_from('<L', self.raw, 0x2C)[0]
        self._load_difat()
        self._load_fat()
        self.dir_chain = self._chain(self.dir_start)
        self.entries = self._load_entries()
        self._load_mini()

    def _sector_off(self, sid: int) -> int:
        return (sid + 1) * self.sector_size

    def _load_difat(self):
        self.fat_sectors = list(struct.unpack_from('<109L', self.raw, 0x4C))
        first = struct.unpack_from('<L', self.raw, 0x44)[0]
        count = struct.unpack_from('<L', self.raw, 0x48)[0]
        sid = first
        per = self.sector_size // 4
        for _ in range(count):
            if sid >= 0xFFFFFFFA:
                break
            off = self._sector_off(sid)
            vals = list(struct.unpack_from('<%dL' % per, self.raw, off))
            self.fat_sectors += vals[:-1]
            sid = vals[-1]
        self.fat_sectors = [s for s in self.fat_sectors[:self.n_fat] if s < 0xFFFFFFFA]

    def _load_fat(self):
        per = self.sector_size // 4
        self.fat = []
        for s in self.fat_sectors:
            off = self._sector_off(s)
            self.fat += list(struct.unpack_from('<%dL' % per, self.raw, off))

    def _chain(self, start: int):
        chain, sid, guard = [], start, 0
        while sid < 0xFFFFFFFA and guard < 100000:
            chain.append(sid)
            sid = self.fat[sid]
            guard += 1
        return chain

    def _load_entries(self):
        blob = bytearray()
        for sid in self.dir_chain:
            off = self._sector_off(sid)
            blob += self.raw[off:off + self.sector_size]
        entries = []
        for i in range(len(blob) // 128):
            rec = blob[i * 128:(i + 1) * 128]
            nlen = struct.unpack_from('<H', rec, 0x40)[0]
            name = rec[:max(0, nlen - 2)].decode('UTF-16LE', 'replace') if nlen >= 2 else ''
            entries.append({
                'index': i,
                'name': name,
                'type': rec[0x42],
                'start': struct.unpack_from('<L', rec, 0x74)[0],
                'size': struct.unpack_from('<Q', rec, 0x78)[0],
            })
        return entries

    def entry(self, name: str):
        for e in self.entries:
            if e['name'] == name and e['type'] == 2:
                return e
        return None

    # ---- 迷你流（<4096 字节的流，存在 root entry 的流里） ----
    def _load_mini(self):
        self.mini_sector = 1 << self.mini_shift
        first = struct.unpack_from('<L', self.raw, 0x3C)[0]
        self.minifat = []
        if first < 0xFFFFFFFA:
            for sid in self._chain(first):
                off = self._sector_off(sid)
                per = self.sector_size // 4
                self.minifat += list(struct.unpack_from('<%dL' % per, self.raw, off))
        root = next((e for e in self.entries if e['type'] == 5), None)
        self.mini_chain = self._chain(root['start']) if root else []
        self.mini_bytes = bytearray()
        for sid in self.mini_chain:
            off = self._sector_off(sid)
            self.mini_bytes += self.raw[off:off + self.sector_size]

    def _mini_write(self, mini_off: int, data: bytes):
        """把数据写回迷你流（镜像到文件里的 root 流扇区）。"""
        pos = mini_off
        done = 0
        while done < len(data):
            sec_idx = pos // self.sector_size
            in_sec = pos % self.sector_size
            sid = self.mini_chain[sec_idx]
            base = self._sector_off(sid) + in_sec
            take = min(self.sector_size - in_sec, len(data) - done)
            chunk = data[done:done + take]
            self.raw[base:base + take] = chunk
            self.mini_bytes[pos:pos + take] = chunk
            done += take
            pos += take

    def read_stream(self, name: str) -> bytes:
        e = self.entry(name)
        if e is None:
            raise KeyError('stream not found: %s' % name)
        if e['size'] < self.mini_cutoff:
            chain = []
            sid, guard = e['start'], 0
            while sid < 0xFFFFFFFA and guard < 100000:
                chain.append(sid)
                sid = self.minifat[sid]
                guard += 1
            buf = bytearray()
            for msid in chain:
                buf += self.mini_bytes[msid * self.mini_sector:(msid + 1) * self.mini_sector]
            return bytes(buf[:e['size']])
        chain = self._chain(e['start'])
        buf = bytearray()
        for sid in chain:
            off = self._sector_off(sid)
            buf += self.raw[off:off + self.sector_size]
        return bytes(buf[:e['size']])

    def capacity(self, name: str) -> int:
        e = self.entry(name)
        if e['size'] < self.mini_cutoff and e['start'] < 0xFFFFFFFA:
            sid, n = e['start'], 0
            while sid < 0xFFFFFFFA and n < 100000:
                n += 1
                sid = self.minifat[sid]
            return n * self.mini_sector
        return len(self._chain(e['start'])) * self.sector_size

    def _patch_entry_size(self, index: int, size: int):
        off = index * 128
        sid_index = off // self.sector_size
        in_off = off % self.sector_size
        sid = self.dir_chain[sid_index]
        base = self._sector_off(sid) + in_off
        struct.pack_into('<Q', self.raw, base + 0x78, size)

    def write_stream_inplace(self, name: str, data: bytes):
        e = self.entry(name)
        if e is None:
            raise KeyError('stream not found: %s' % name)
        cap = self.capacity(name)
        if len(data) > cap:
            raise ValueError('%s: 新内容 %d 字节 > 原有空间 %d 字节' % (name, len(data), cap))
        if e['size'] < self.mini_cutoff:
            self._mini_write(e['start'] * self.mini_sector, data)
            self._patch_entry_size(e['index'], len(data))
            e['size'] = len(data)
            return
        chain = self._chain(e['start'])
        pos = 0
        for sid in chain:
            off = self._sector_off(sid)
            take = min(self.sector_size, len(data) - pos)
            chunk = data[pos:pos + take]
            self.raw[off:off + self.sector_size] = chunk.ljust(self.sector_size, b'\x00')
            pos += take
            if pos >= len(data):
                break
        self._patch_entry_size(e['index'], len(data))
        e['size'] = len(data)

    def save(self, path: str):
        with open(path, 'wb') as f:
            f.write(self.raw)


# --------------------------------------------------------------------------
# 模块名 <-> 流名
# --------------------------------------------------------------------------
CODE_PAGE = 'cp936'      # 本工程 PROJECTCODEPAGE = 936（简体中文）


def module_name_of(path: str) -> str:
    base = os.path.basename(path)
    return os.path.splitext(base)[0]


def load_sources(src_dir: str):
    """读仓库里的 UTF-8 源码，返回 {模块名: 字节流(cp936, CRLF)}"""
    out = {}
    for fn in sorted(os.listdir(src_dir)):
        if not fn.lower().endswith(('.bas', '.cls')):
            continue
        text = open(os.path.join(src_dir, fn), encoding='utf-8').read()
        if text.startswith('\ufeff'):
            text = text[1:]
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        text = text.replace('\n', '\r\n')
        if not text.endswith('\r\n'):
            text += '\r\n'
        out[module_name_of(fn)] = text.encode(CODE_PAGE)
    return out


def extract(pptm: str, src_dir: str):
    os.makedirs(src_dir, exist_ok=True)
    z = zipfile.ZipFile(pptm)
    cfb = Cfb(z.read('ppt/vbaProject.bin'))
    written = []
    for e in cfb.entries:
        if e['type'] != 2:
            continue
        name = e['name']
        if not re.match(r'^(模块\d+|Module\d+|Slide\d+|Sheet\d+|ThisWorkbook|ThisPresentation)$', name):
            continue
        raw = cfb.read_stream(name)
        if raw[:1] != b'\x01':
            continue
        code = ovba_decompress(raw)
        # 只保留源码部分（p-code 可能在源码之后的字节里，这里按 Attribute 头截断）
        text = code.decode(CODE_PAGE, 'replace')
        if 'Attribute VB_Name' in text:
            idx = text.find('Attribute VB_Name')
            text = text[idx:]
        ext = '.cls' if name.startswith(('Slide', 'Sheet', 'This')) else '.bas'
        path = os.path.join(src_dir, name + ext)
        with open(path, 'w', encoding='utf-8', newline='') as f:
            f.write(text)
        written.append((name, len(text), e['size']))
    print('抽取 %d 个模块 -> %s' % (len(written), src_dir))
    for name, n, orig in written:
        print('  %-12s 源码 %6d 字符  原流 %6d 字节' % (name, n, orig))


def build(src_dir: str, in_pptm: str, out_pptm: str, keep_srp=False):
    sources = load_sources(src_dir)
    if not sources:
        raise SystemExit('src 目录里没有 .bas/.cls')
    zin = zipfile.ZipFile(in_pptm)
    vba = zin.read('ppt/vbaProject.bin')
    cfb = Cfb(vba)

    patched = []
    for name, code in sources.items():
        ent = cfb.entry(name)
        if ent is None:
            print('  跳过（pptm 里没有该模块）: %s' % name)
            continue
        is_regular = ent['size'] >= cfb.mini_cutoff
        container = ovba_compress(code)
        mode = 'LZ'
        # CFB 规则：< 4096 字节的流必须放进迷你流。原本是普通流的模块不能缩到 4096 以下，
        # 否则读取方会按迷你流解析 → 乱码。此时改用压缩率更差的字面量编码把流撑住。
        if is_regular and len(container) < cfb.mini_cutoff:
            lit = ovba_compress_literal(code)
            if len(lit) >= cfb.mini_cutoff:
                container, mode = lit, '字面量'
            else:
                raise SystemExit('%s 压缩后 %d 字节、字面量 %d 字节，都低于迷你流阈值 4096；'
                                 '请在源码末尾补一段注释填充' % (name, len(container), len(lit)))
        if (not is_regular) and len(container) >= cfb.mini_cutoff:
            raise SystemExit('%s 原本是迷你流，现在超过 4096 字节，需要改成普通流分配' % name)
        back = ovba_decompress(container)
        assert back == code, '%s 压缩/解压往返失败' % name
        cap = cfb.capacity(name)
        if len(container) > cap:
            raise SystemExit('%s 编码后 %d 字节，超出原空间 %d 字节' % (name, len(container), cap))
        cfb.write_stream_inplace(name, container)
        patched.append((name, len(code), len(container), cap, mode))

    if not keep_srp:
        for e in cfb.entries:
            if e['type'] == 2 and e['name'].startswith('__SRP_'):
                cfb._patch_entry_size(e['index'], 0)
                e['size'] = 0

    tmp_vba = out_pptm + '.vbaProject.bin'
    cfb.save(tmp_vba)

    with zipfile.ZipFile(out_pptm, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'ppt/vbaProject.bin':
                data = open(tmp_vba, 'rb').read()
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zout.writestr(zi, data)
    os.remove(tmp_vba)

    print('重建完成 -> %s' % out_pptm)
    for name, nsrc, nenc, cap, mode in patched:
        print('  %-12s 源码 %6d → 流 %6d 字节（%s / 原空间 %6d）' % (name, nsrc, nenc, mode, cap))
    if not keep_srp:
        print('  已清空 __SRP_* 性能缓存（强制宿主重新编译）')


def verify(pptm: str, src_dir: str):
    sources = load_sources(src_dir)
    z = zipfile.ZipFile(pptm)
    cfb = Cfb(z.read('ppt/vbaProject.bin'))
    bad = 0
    for name, code in sources.items():
        e = cfb.entry(name)
        if e is None:
            print('  [缺失] %s' % name)
            bad += 1
            continue
        got = ovba_decompress(cfb.read_stream(name))
        text = got.decode(CODE_PAGE, 'replace')
        exp = code.decode(CODE_PAGE, 'replace')
        idx = text.find('Attribute VB_Name')
        if idx > 0:
            text = text[idx:]
        if text.rstrip('\x00\r\n') == exp.rstrip('\r\n'):
            print('  [一致] %-12s %d 字节' % (name, len(code)))
        else:
            print('  [不一致] %s' % name)
            for i, (a, b) in enumerate(zip(text.split('\r\n'), exp.split('\r\n'))):
                if a != b:
                    print('      第 %d 行：' % (i + 1))
                    print('        pptm: %s' % a[:90])
                    print('        源码: %s' % b[:90])
                    break
            bad += 1
    return 1 if bad else 0


def lint(src_dir: str):
    """粗检：Sub/Function 块平衡、If/For/While 平衡、行内引号成对。"""
    problems = 0
    for fn in sorted(os.listdir(src_dir)):
        if not fn.lower().endswith(('.bas', '.cls')):
            continue
        path = os.path.join(src_dir, fn)
        lines = open(path, encoding='utf-8').read().splitlines()
        stack = []
        in_proc = False
        for no, raw in enumerate(lines, 1):
            line = raw.strip()
            if line.startswith("'"):
                continue
            code = line.split("'")[0] if line.count('"') % 2 == 0 else line
            if code.count('"') % 2 == 1:
                print('%s:%d 引号不成对: %s' % (fn, no, line[:80]))
                problems += 1
            low = code.lower()
            if re.match(r'^(public |private |friend )?(sub|function|property (get|let|set)) ', low):
                in_proc = True
                stack.append((low.split()[0], no))
                continue
            if re.match(r'^end (sub|function|property)\b', low):
                if not stack:
                    print('%s:%d 多余的 End: %s' % (fn, no, line[:60]))
                    problems += 1
                else:
                    stack.pop()
                continue
            if re.match(r'^(if\b.*\bthen|for\b|do\b|while\b|with\b|select case\b)', low) and not re.search(r'\bthen\b.*[^\s]', low) is False:
                pass
            if re.match(r'^if\b', low) and re.search(r'\bthen\b', low) and not re.search(r'\bthen\b\s*\S', low):
                stack.append(('if', no))
            elif re.match(r'^for\b', low) or re.match(r'^do\b', low) or re.match(r'^with\b', low) or re.match(r'^select case\b', low):
                stack.append(('for', no))
            elif re.match(r'^end if\b', low) or re.match(r'^next\b', low) or re.match(r'^loop\b', low) or re.match(r'^end with\b', low) or re.match(r'^end select\b', low):
                if stack:
                    stack.pop()
                else:
                    print('%s:%d 多余的块结束: %s' % (fn, no, line[:60]))
                    problems += 1
        for kind, no in stack:
            print('%s:%d 未闭合的块（%s）' % (fn, no, kind))
            problems += 1
        print('  检查完毕: %s（%d 行，%d 处疑似问题）' % (fn, len(lines), problems))
    return 0


def keys():
    raw = {
        'data': 'ak_8a63516df24a4da9ed9e4be9',
        'ai': 'ak_aaa6a0be6c2c7793c200d49f',
        'calc': 'ak_db3340f199bbe16354f85d11',
        'login': '53f2d9e16285159fd50f78dc757a8d1b',
    }
    for k, v in raw.items():
        enc = ''.join('%02X' % (b ^ 90) for b in v.encode('ascii'))
        print('%-6s %s -> %s' % (k, v, enc))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)
    p = sub.add_parser('extract'); p.add_argument('pptm'); p.add_argument('src_dir')
    p = sub.add_parser('build'); p.add_argument('src_dir'); p.add_argument('in_pptm'); p.add_argument('out_pptm'); p.add_argument('--keep-srp', action='store_true')
    p = sub.add_parser('verify'); p.add_argument('pptm'); p.add_argument('src_dir')
    p = sub.add_parser('lint'); p.add_argument('src_dir')
    sub.add_parser('keys')
    a = ap.parse_args()
    if a.cmd == 'extract':
        extract(a.pptm, a.src_dir)
    elif a.cmd == 'build':
        build(a.src_dir, a.in_pptm, a.out_pptm, a.keep_srp)
    elif a.cmd == 'verify':
        sys.exit(verify(a.pptm, a.src_dir))
    elif a.cmd == 'lint':
        sys.exit(lint(a.src_dir))
    elif a.cmd == 'keys':
        keys()


if __name__ == '__main__':
    main()
