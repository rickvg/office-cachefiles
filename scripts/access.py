# env = python3.5
# Compatible with: Office 2016
import sys
import binascii
import struct
import datetime
from io import BytesIO
import os


def convert_type_to_size(col_type, col_len):  # Converts column types to bitsize
    type_size = {
        "01": 0,
        "02": 8,
        "03": 16,
        "04": 32,
        "05": 64,
        "06": 32,
        "07": 32,
        "08": 64,
        "09": 2040,
        "0A": 2040,
        "0B": 96,
        "0C": 96,
        "0D": "UNKNOWN",
        "0E": "UNKNOWN",
        "0F": 128,
        "10": 136
    }
    col_type = str(col_type).upper()
    if col_type == "0A" or col_type == "09":
        return col_len*8

    return type_size[col_type]


def unpack_data_to_type(input_data, col_type):
    type_fmt = {
        "02": "B",
        "03": "h",
        "04": "i",
        "05": "Q",
        "06": "f",
        "07": "d",
    }
    if col_type not in type_fmt:
        if col_type.upper() == "0F":  # Format to GUID-string
            data1 = (input_data[0:4])[::-1]
            data2 = (input_data[4:6])[::-1]
            data3 = (input_data[6:8])[::-1]
            data4a = input_data[8:10]
            data4b = input_data[10:]
            components = [data1, data2, data3, data4a, data4b]
            hex_bytes = [binascii.hexlify(d) for d in components]
            decoded = [bytes.decode(b) for b in hex_bytes]
            return '-'.join(decoded)

        elif col_type.upper() == "08":  # Convert FILETIME to human-readable format, also OLETIME
            hex_value = binascii.hexlify(input_data)
            print(hex_value)
            dt = hex_value
            noerror = 0
            try:
                (s, ns100) = divmod(int(hex_value, 16) - 116444726000000000, 10000000)
                dt = datetime.datetime.utcfromtimestamp(s)
                dt = dt.replace(microsecond=(ns100 // 10))
                dt = dt.strftime("%d/%M/%Y %H:%M:%S")
                noerror = 1
            except:
                print("Not FILETIME")

            try: # To Fix: OLETIME conversion
                newval = struct.unpack('d', input_data)[0]
                OLE_TIME_ZERO = datetime.datetime(1899, 12, 30, 0, 0, 0)
                dt = OLE_TIME_ZERO + datetime.timedelta(days=float(newval))
                dt = dt.strftime("%d/%M/%Y %H:%M:%S")
                noerror = 1
            except:
                print("Not OLETIME")

            if noerror == 0:
                return binascii.hexlify(input_data)

            return dt

        elif col_type.upper() == "0C": # Read memo data and put in array
            mem_len = struct.unpack('<i', input_data[:3] + '\x00')[0]
            mem_bitmask = hex(struct.unpack('<B', input_data[3:4])[0])
            mem_pointer = struct.unpack('<i', input_data[4:8])[0]
            return [mem_len, mem_bitmask, mem_pointer]

        return binascii.hexlify(input_data)
    else:
        return struct.unpack(type_fmt[col_type], input_data)[0]


PAGE_SIZE = 4096  # Size in bytes
PAGE_TYPE = 0x02  # TableDef Page

if len(sys.argv) != 2:
    print("Usage: python access.py [access database]")
    exit()

inp_file = open(sys.argv[1], "rb")
folder_name = os.path.dirname(sys.argv[1])
file_contents = inp_file.read()
file_size = len(file_contents)
amount_of_pages = len(file_contents)//PAGE_SIZE

table_pointer_columns = {}
table_var_cols = {}


def retrieve_tabledef(table_pointers, table_var_col, table_offset_pg, allow_large_data = True):
    offset = hex(inp_file.tell() - 1)
    inp_file.seek(inp_file.tell()-1)
    print(hex(inp_file.tell()))
    prev_offset = offset

    page_data = inp_file.read(4096)
    print("Offset: " + str(offset))

    next_pg = struct.unpack("i", page_data[4:8])[0]
    print("Next page: " + str(next_pg))
    if next_pg != 0:
        pagecounter = 2
        pages = 1
        while pages < pagecounter:
            print(pagecounter)
            print(next_pg)
            table_offset_pg[offset] = hex(next_pg) + "000"
            inp_file.seek(int(table_offset_pg[offset], 16))
            offset = hex(inp_file.tell())
            oldpos = inp_file.tell()
            inp_file.read(8)
            page_data = page_data + inp_file.read(4088)
            inp_file.seek(oldpos)
            inp_file.read(4)
            next_pg = struct.unpack("i", inp_file.read(4))[0]
            if next_pg != 0:
                pagecounter += 1
            pages += 1

    offset = prev_offset
    page_data = BytesIO(page_data)
    page_data.read(8)
    length_data = struct.unpack("i", page_data.read(4))[0]

    if offset not in table_pointers:
        table_pointers[offset] = []
    if offset not in table_var_col:
        table_var_col[offset] = []

    print("Length of data:" + str(length_data))
    page_data.read(31)
    num_var_cols = struct.unpack("H", page_data.read(2))[0]
    print("Num_var_cols:" + str(num_var_cols))

    if length_data < 4096 or allow_large_data:  # Length of data can be larger if next_pg is set.
        table_var_col[offset].append(num_var_cols)
        num_cols = struct.unpack("H", page_data.read(2))[0]
        print("Num_columns:" + str(num_cols))
        num_idx = struct.unpack("i", page_data.read(4))[0]
        print("Num_idx:" + str(num_idx))
        num_real_idx = struct.unpack("i", page_data.read(4))[0]
        print("Num_real_idx:" + str(num_real_idx))
        ptr_used_pages = struct.unpack("i", page_data.read(4))[0]
        inp_file.seek(ptr_used_pages)
        map = struct.unpack("i", inp_file.read(4))
        print(map)

        page_data.read(4)

        for j in range(0, num_real_idx):
            page_data.read(12)

        column_types = []
        for j in range(0, num_cols):
            print(page_data.tell())
            column_type = binascii.hexlify(page_data.read(1))
            print("Column Type:" + str(column_type))
            page_data.read(4)
            var_col_num = struct.unpack("H", page_data.read(2))[0]
            print("Offset variable length columns:" + str(struct.unpack("H", page_data.read(2))[0]))
            column_number = struct.unpack("H", page_data.read(2))[0]
            print("Column Number:" + str(column_number))
            page_data.read(4)
            column_flag = str(binascii.hexlify(page_data.read(1)))
            print("Flag: " + column_flag)
            page_data.read(5)
            offset_fixed = struct.unpack("H", page_data.read(2))[0]
            print("Offset fixed length columns:" + str(offset_fixed))
            col_length = struct.unpack("H", page_data.read(2))[0]
            print("Col Length:" + str(col_length))
            toappend = str(column_type) + ":" + str(col_length) + ":" + str(column_flag) + ":" + str(
                    offset_fixed) + ":" + str(column_number) + ":" + str(var_col_num)
            column_types.append(toappend)

        column_names = []
        for j in range(0, num_cols):
            col_name_length = struct.unpack("H", page_data.read(2))[0]
            column_name = str(page_data.read(col_name_length).decode("utf-16"))
            column_names.append(column_name)
            print("Column Name:" + str(column_name))

        print(len(column_names))
        print(len(column_types))
        columns_names_types = []
        for j in range(0, len(column_names)):
            columns_names_types.append(str(column_names[j]) + ":" + str(column_types[j]))

        table_pointers[str(offset)].append(columns_names_types)

        for j in range(0, num_real_idx):
            page_data.read(4)

        for j in range(0, 10):
            page_data.read(21)

        for j in range(0, num_idx):
            page_data.read(24)

        for j in range(0, num_idx):
            idx_name_length = struct.unpack("H", page_data.read(2))[0]
            #print("IDX NAME: " + str(page_data.read(idx_name_length)))


        #print(page_data, table_pointers, table_var_col, table_offset_pg)
        #print(table_pointers)
    return page_data, table_pointers, table_var_col, table_offset_pg


table_offset_nextpg = {}  # Contains offset of table containing next_pg value
page_dataset = {}

# Start reading the Table definitions
for i in range(0, amount_of_pages):
    inp_file.seek((i * PAGE_SIZE))
    offst = hex(inp_file.tell())
    _type = binascii.hexlify(inp_file.read(1))
    if _type == "02" and hex(inp_file.tell()-1) not in table_offset_nextpg.values():
        try:
            page_data, table_pointer_columns, table_var_cols, table_offset_nextpg = retrieve_tabledef(table_pointer_columns, table_var_cols, table_offset_nextpg)
            page_dataset[offst] = page_data
        except:
            print("Error on " + str(i))

table_pointer_rows = {}  # Contains the row offsets based on the datapage pointer
table_pointer_datapage = {}  # Contains the tabledef offsets based on the datapage pointer
print("")
print(table_pointer_columns)

# Retrieve row offsets from datapages
for i in range(0, amount_of_pages):
    inp_file.seek((i * PAGE_SIZE))
    _type = binascii.hexlify(inp_file.read(1))
    inp_file.read(3)
    value = inp_file.read(4)
    inp_file.seek(inp_file.tell() - 7)
    if value == "LVAL":
        print("LVAL")
        print(hex(inp_file.tell()))

    if _type == "01" and value != "LVAL": #Data Page
        offset = inp_file.tell()-1
        print("Offset: " + str(offset))
        print("Unknown:" + str(struct.unpack("B", inp_file.read(1))[0]))
        print("Free Space:" + str(struct.unpack("H", inp_file.read(2))[0]))
        tdef_pg = hex(struct.unpack(">i", inp_file.read(4))[0])

        table_pointer_datapage[offset] = tdef_pg

        print("TDEF_PG:" + str(tdef_pg))
        print("Unknown:" + str(struct.unpack("i", inp_file.read(4))[0]))
        num_rows = struct.unpack("<H", inp_file.read(2))[0]

        if tdef_pg not in table_pointer_rows:
            table_pointer_rows[offset] = []

        row_offsets = []
        print("Number of Rows:" + str(num_rows))
        for j in range(0, num_rows):
            row_offset = struct.unpack("<H", inp_file.read(2))[0]
            row_offsets.append(row_offset)
            print("Row offset:" + str(row_offset))
        table_pointer_rows[offset].append(row_offsets)
    print("")

# table_pointer_rows Contains the row offsets based on the datapage pointer
# table_pointer_datapage Contains the tabledef offsets based on the datapage pointer
# table_pointer_columns Contains the column info based on the tabledef pointer

print(table_pointer_rows)

print(table_pointer_datapage)
print(len(table_pointer_datapage))
counter = 0

tablerows = {}
tablecolumns = {}

# Retrieve rows from datapages and link them to columns
for entry in table_pointer_datapage:
    tabledef_pointer = str(table_pointer_datapage[entry])[:-3]
    if tabledef_pointer in table_pointer_columns:  # All 0x0 are filtered out here
        counter += 1
        datapage_pointer = entry
        print("Datapage location: " + str(hex(datapage_pointer)))
        print("TableDEF location: " + str(hex(int(tabledef_pointer, 16))))
        try:
            for k in range(0, len(table_var_cols[tabledef_pointer])):
                for i in range(0, len(table_pointer_rows[entry])):
                    prev_pointer = PAGE_SIZE - 1
                    for j in table_pointer_rows[entry][i]:  # j = row offset. but sometimes the row offset is too large.
                        if j > 4096:
                            old_position = 0
                            new_position = 0
                            page_counter = 1
                            while j < old_position or j > new_position:
                                old_position = new_position
                                new_position = 4096 * page_counter
                                page_counter += 1

                            amount_of_pages = page_counter-1
                            j = 4096 - (new_position - j)
                            print("PAGES: " + str(amount_of_pages))
                            print("NEW OFFSET: " + str(j))
                        position = datapage_pointer + j
                        toread = prev_pointer - j
                        inp_file.seek(position)
                        row_cols = struct.unpack("H", inp_file.read(2))[0]
                        bitmask = (row_cols + 7) / 8
                        oldpos = inp_file.tell()
                        inp_file.seek(datapage_pointer + prev_pointer - bitmask + 1)
                        nullmask = inp_file.read(bitmask)  # Bitwise AND operation: & in C -> nullmask[byte_num] & (1 << bit_num)
                        print("NULLMASK: " + str(nullmask))
                        inp_file.seek(oldpos)

                        if table_var_cols[tabledef_pointer][k] >= 0:
                            newpos = datapage_pointer + prev_pointer - bitmask - 1
                            oldpos = inp_file.tell()

                            inp_file.seek(newpos)
                            print("ROW OFFSETS NOWW: " + str(inp_file.tell()))
                            row_var_cols = struct.unpack("H", inp_file.read(2))[0]
                            print("Row cols: " + str(row_cols))
                            print("Row var cols: " + str(row_var_cols))
                            print("Actual var cols:" + str(table_var_cols[tabledef_pointer]))
                            if row_var_cols > table_var_cols[tabledef_pointer][k]:
                                exit()
                            inp_file.seek(oldpos)

                            var_col_offsets = []

                            # Retrieve column offsets
                            for q in range(0, row_var_cols + 1):
                                oldpos = inp_file.tell()
                                offset = prev_pointer - bitmask - 3 - (q * 2)
                                inp_file.seek(datapage_pointer + offset)
                                var_col_offsets.append(struct.unpack("H", inp_file.read(2))[0])
                                # print(prev_pointer - bitmask - 3 - (q * 2))
                                inp_file.seek(oldpos)

                            row_fixed_cols = row_cols - row_var_cols

                            oldpos = inp_file.tell()
                            inp_file.seek(position)
                            #print("Offset: " + str(inp_file.tell()))
                            var_col_counter = 0
                            row_data = []
                            column_data = []
                            for q in range(0, len(table_pointer_columns[tabledef_pointer][k])):
                                column_number = q
                                column = table_pointer_columns[tabledef_pointer][k][q].split(":")[0]
                                if tabledef_pointer not in tablecolumns:
                                    tablecolumns[tabledef_pointer] = []
                                    tablerows[tabledef_pointer] = []

                                column_type = table_pointer_columns[tabledef_pointer][k][q].split(":")[1]
                                column_length = int(table_pointer_columns[tabledef_pointer][k][q].split(":")[2])
                                column_flag = table_pointer_columns[tabledef_pointer][k][q].split(":")[3]
                                fixed_offset = table_pointer_columns[tabledef_pointer][k][q].split(":")[4]
                                col_nr = table_pointer_columns[tabledef_pointer][k][q].split(":")[5]
                                var_col_nr = table_pointer_columns[tabledef_pointer][k][q].split(":")[6]

                                print("Offset: " + str(inp_file.tell()))
                                print(column)
                                if column_length == 0:
                                    column_length = 12
                                print("Col flag: " + str(column_flag))
                                print("Col Type: " + str(column_type))
                                if column_flag == "02":
                                    fixed = 0
                                else:
                                    fixed = 1

                                number_of_columns = len(table_pointer_columns[tabledef_pointer][k])
                                byte_num = int(var_col_nr) / 8
                                bit_num = int(var_col_nr) % 8

                                print("BIT NUM: " + str(1 << bit_num))
                                print(ord(nullmask[byte_num]))
                                print(ord(nullmask[byte_num]) & (1 << bit_num))
                                # Check nullmask whether column = null
                                # Result is inverse
                                if ord(nullmask[byte_num]) & (1 << bit_num):
                                    isnull = False
                                else:
                                    isnull = True
                                fixed_cols_found = 0
                                print(isnull)

                                # Read data from rows: Either fixed or variable
                                if fixed == 1 and fixed_cols_found < row_fixed_cols:
                                    print("FIXED OFFSET: " + str(fixed_offset))
                                    col_start = int(fixed_offset) + 2  # To add: fixed offset
                                    start = j + col_start + datapage_pointer
                                    oldpos = inp_file.tell()
                                    inp_file.seek(start)
                                    data = inp_file.read(convert_type_to_size(column_type, column_length)/8)
                                    if isnull == 0:
                                        field_value = unpack_data_to_type(data, column_type)
                                    else:
                                        field_value = ""
                                    fixed_cols_found += 1
                                elif fixed == 0:
                                    col_start = var_col_offsets[var_col_counter]
                                    start = j + col_start + datapage_pointer
                                    size = var_col_offsets[var_col_counter+1] - col_start
                                    if var_col_offsets[var_col_counter] == var_col_offsets[var_col_counter+1]:
                                        isnull = 1

                                    inp_file.seek(start)
                                    data = inp_file.read(convert_type_to_size(column_type, column_length)/8)
                                    var_col_counter += 1
                                    if isnull == 0:
                                        field_value = unpack_data_to_type(data, column_type)
                                    else:
                                        field_value = ""

                                if str(column_type).upper() == "0C" and isnull == 0:

                                    print(field_value[1])
                                    if field_value[1] == '0x80':
                                        field_value = inp_file.read(field_value[0])
                                        try:
                                            field_value = field_value.decode('utf-16')
                                        except:
                                            field_value = binascii.hexlify(field_value)
                                    elif field_value[1] == '0x40':
                                        field_value = "LVAL record 1"
                                    elif field_value[1] == '0x00':
                                        field_value = "LVAL record 2"
                                try:
                                    print("DATA: " + str(field_value))
                                except:
                                    print("DATA: Not printable")
                                row_data.append(field_value)
                                column_data.append(column)
                            tablerows[tabledef_pointer].append(row_data)
                            tablecolumns[tabledef_pointer] = column_data
                        else:
                            row_fixed_cols = row_cols
                            print("Only fixed columns: " + row_fixed_cols)
                        prev_pointer = j - 1
        except:
            print("DIDN'T WORK.")
print("Lost values: " + str(len(table_pointer_datapage)-counter))
#print(table_pointer_datapage)
#print(table_offset_nextpg)
#print(table)

for i in tablecolumns:
    filename = folder_name + "/Table_" + i + ".csv"
    first_line = ""
    second_line = ""
    rows = []
    csv_file = open(filename, "w")
    for j in tablecolumns[i]:
        j = str(j)
        j = j.replace('\x00', '')
        first_line = first_line + "," + j
    csv_file.write(first_line + "\n")
    print(first_line)
    for j in tablerows[i]:
        second_line = ""
        for p in j:
            p = str(p)
            p = p.replace('\x00', '')
            second_line = second_line + "," + str(p)
        print(second_line)
        csv_file.write(second_line + "\n")
    csv_file.close()

for i in table:
    print(table[i])
    print("")
#print(table_pointer_columns)