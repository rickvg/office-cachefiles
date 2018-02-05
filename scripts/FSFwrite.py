import sys
import os
import binascii

if len(sys.argv) != 3:
    print("Usage: FSFwrite.py [FSD-file] [FSF-filename]")
    exit()

if not os.path.isfile(sys.argv[1]):
    print("Not a file!")
    exit()

fsfname = str(sys.argv[2]) + ".FSF"
fsffile = open(fsfname, "wb")
fsdname = sys.argv[1]

header = b"\xB3\xE2\x48\xEA\xA6\xA1\xDD\x40\xBC\x06\xC7\xC4\x62\x92\x12\x71\x0C\x00\x18\xBA"
length = b"\x5D"
data = fsdname.encode("utf-16")[2:]
length = binascii.unhexlify(format(len(data)+1, '02x'))
end = b"\x05"

filecontent = header + length + data + end
fsffile.write(filecontent)
