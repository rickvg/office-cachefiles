import re

fileName = "FSD-{89C56A54-B755-4EA7-AAC6-37189248FCE2}.FSD"
f = open(fileName, "rb")
content = f.read()

# Check for FSD-file format based on the Magic value
if content[0:16] != b'\x0c\x83\xd2\x91\xae\x1b\xd4\x4d\xaa\x65\x46\x79\xfb\xda\xdd\x7a':
    print("The magic value for FSD-format is incorrect, please input an FSD-file (Tested with MS Office 2016)")
    exit()

pkHeaderAddresses = []
delimeterAddresses = []
delimeter2Addresses = []

# Find positions of PK, I and A headers in file
for m in re.finditer(b'\x50\x4b\x03\x04', content):
    pkHeaderAddresses.append(m.start())

for m in re.finditer(b'\x50\x4b\x01\x02', content):
    pkHeaderAddresses.append(m.start())

for m in re.finditer(b'\x50\x4b\x07\x08', content):
    pkHeaderAddresses.append(m.start())

for m in re.finditer(b'\xcf\xaa\x69\x49', content):
    delimeterAddresses.append(m.start())

for m in re.finditer(b'\xc4\xf4\xf7\xf5', content):
    delimeter2Addresses.append(m.start())

pkHeaderAddresses = sorted(pkHeaderAddresses)
delimeter2Addresses = sorted(delimeter2Addresses)
delimeterAddresses = sorted(delimeterAddresses)

# Number of PK headers
print("Number of PK Headers", len(pkHeaderAddresses))

# Number of Delimeter: ÏªiIÿ
print('Number of I Headers', len(delimeterAddresses))

# Number of Delimeter: Äô÷õ
print('Number of A Headers', len(delimeter2Addresses))
print()

print("PK header")
print(pkHeaderAddresses)
print()
print("Äô÷õ header")
print(delimeter2Addresses)
print()
print("ÏªiIÿ header")
print(delimeterAddresses)
print()

aHeaderFinalAddress = 0;
for i in range(0, len(delimeter2Addresses) - 1):
    if delimeter2Addresses[i] > pkHeaderAddresses[len(pkHeaderAddresses) - 1]:
        aHeaderFinalAddress = delimeter2Addresses[i]
        break

print("Final required A header ", aHeaderFinalAddress)
print()

new_pk_header = []
new_i_header = []

# Find next higher I-header start position from PK perspective and put it in a list
# The I-header list is never longer than pkHeader length
for pk_address in pkHeaderAddresses:
    new_pk_header.append(pk_address)
    for i_address in delimeterAddresses:
        if pk_address < i_address:
            if i_address not in new_i_header:
                new_i_header.append(i_address)
            break

# Check whether A-header comes closer than new I-header
if aHeaderFinalAddress - new_pk_header[len(new_pk_header)-1] < new_i_header[len(new_i_header)-1] - new_pk_header[len(new_pk_header)-1]:
    new_i_header[len(new_i_header)-1] = aHeaderFinalAddress
    print("A-Header is closer to PK than I-header")

print("Cleaned PK Header: " + str(new_pk_header))
print("Cleaned I Header: " + str(new_i_header))

new_data = []
# Read PK until I-header - For last one PK until A
for i in range(0, len(new_i_header)):
    f.seek(new_pk_header[i])
    new_data.append(f.read(new_i_header[i]-new_pk_header[i]))

f.close()

# Write to output file and read I-header section until \x79\x05 = I-header end
f = open("test.docx", "wb")
for data in new_data:
    for m in re.finditer(b'\x79\x05', data[len(data)-16:]):
        data = data[:len(data)-16+m.start()]
    f.write(data)
f.close()
