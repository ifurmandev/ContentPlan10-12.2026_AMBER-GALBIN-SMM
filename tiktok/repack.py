# -*- coding: utf-8 -*-
"""Перепаковує .xlsx максимально щільно (zopfli, без docProps і theme), щоб файл легше пройшов через завантаження."""
import re
import struct
import sys
import zipfile
import zlib

import zopfli.zopfli as zz


def deflate(data):
    # zopfli віддає zlib-потік: відрізаємо 2-байтовий заголовок і 4-байтову контрольну суму
    return zz.compress(data, numiterations=30, gzip_mode=0)[2:-4]


def repack(src, dst):
    with zipfile.ZipFile(src) as z:
        items = {n: z.read(n) for n in z.namelist()}
    for n in list(items):
        if n.startswith("docProps/") or n.startswith("xl/theme/"):
            del items[n]
    ct = items["[Content_Types].xml"].decode()
    ct = re.sub(r'<Override PartName="/(docProps|xl/theme)/[^"]*"[^>]*/>', "", ct)
    items["[Content_Types].xml"] = ct.encode()
    items["_rels/.rels"] = re.sub(rb"<Relationship[^>]*docProps[^>]*/>", b"", items["_rels/.rels"])
    items["xl/_rels/workbook.xml.rels"] = re.sub(rb"<Relationship[^>]*theme[^>]*/>", b"", items["xl/_rels/workbook.xml.rels"])
    order = ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"]
    names = order + sorted(n for n in items if n not in order)
    out, central, off = bytearray(), bytearray(), 0
    for n in names:
        data, fn = items[n], n.encode()
        comp = deflate(data)
        crc = zlib.crc32(data) & 0xFFFFFFFF
        hdr = struct.pack("<IHHHHHIIIHH", 0x04034B50, 20, 0, 8, 0, 0x21, crc, len(comp), len(data), len(fn), 0)
        out += hdr + fn + comp
        central += struct.pack("<IHHHHHHIIIHHHHHII", 0x02014B50, 20, 20, 0, 8, 0, 0x21, crc, len(comp), len(data),
                               len(fn), 0, 0, 0, 0, 0, off) + fn
        off = len(out)
    end = struct.pack("<IHHHHIIH", 0x06054B50, 0, 0, len(names), len(names), len(central), len(out), 0)
    open(dst, "wb").write(bytes(out + central + end))


if __name__ == "__main__":
    repack(sys.argv[1], sys.argv[2])
