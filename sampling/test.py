import struct
import openpyxl as xl


class Wave_Header:
    WAV_HEADER_SIZE = 44
    def __init__(self, file, mode):
        self.file = open('data/' + file + ".wav", mode=mode)
        self.header = self.file.read(Wave_Header.WAV_HEADER_SIZE)  # 44 바이트 헤더
        self.pointer = 0
        self.chunk_id = self.read_4()
        self.chunk_size = self.read_4()
        self.format = self.read_4()
        self.subchunk_id = self.read_4()
        self.subchunk_size = self.read_4()
        self.audio_format = self.read_2()
        self.num_channels = self.read_2()
        self.sample_rate = self.read_4()
        self.byte_rate = self.read_4()
        self.block_align = self.read_2()
        self.bits_per_sample = self.read_2()
        self.subchunk2_id = self.read_4()
        self.subchunk2_size = self.read_4()

    def read_4(self):  # 헤더 4바이트 read
        # print(self.header[self.pointer: self.pointer + 4])
        val = struct.unpack('<i', self.header[self.pointer: self.pointer + 4])  # 4바이트 Little Endian
        self.pointer += 4
        return val[0]

    def read_2(self):  # 헤더 2바이트 read
        # print(self.header[self.pointer: self.pointer + 2])
        val = struct.unpack('<h', self.header[self.pointer: self.pointer + 2])  # 2바이트 Little Endian
        self.pointer += 2
        return val[0]

    def show_info(self):
        print("%-15s : %11s" % ("Chunk ID", self.chunk_id))
        print("%-15s : %11d" % ("Chunk Size", self.chunk_size))
        print("%-15s : %11s" % ("Format", self.format))
        print("%-15s : %11s" % ("Subchunk ID", self.subchunk_id))
        print("%-15s : %11d" % ("Subchunk Size", self.subchunk_size))
        print("%-15s : %11d" % ("Audio Format", self.audio_format))
        print("%-15s : %11d" % ("Num Channels", self.num_channels))
        print("%-15s : %11d" % ("Sample Rate", self.sample_rate))
        print("%-15s : %11d" % ("Byte Rate", self.byte_rate))
        print("%-15s : %11d" % ("Block Align", self.block_align))
        print("%-15s : %11d" % ("Bits Per Sample", self.bits_per_sample))
        print("%-15s : %11s" % ("Subchunk2 ID", self.subchunk2_id))
        print("%-15s : %11d" % ("Subchunk2 Size", self.subchunk2_size))


class Wave(Wave_Header):
    SIZE_SHORT = 2

    def __init__(self, file, mode):
        super().__init__(file, mode)
        self.file_name = file
        self.data = self.file.read()                        # 데이터
        self.n_sample = self.file.tell() // Wave.SIZE_SHORT # 샘플링
        self.decoded_data = list(struct.iter_unpack('h', self.data))

    def close(self):
        self.file.close()

    def show_data(self):
        for i in range(len(self.decoded_data)):
            print(self.decoded_data[i][0], end="\t")
            if i % 40 == 0:
                print()

    def make_exel(self):
        try:
            wb = xl.load_workbook(self.file_name + '.xlsx')
        except:
            wb = xl.Workbook()
        sheet = wb.active
        cell = 'A'
        for i in range(len(self.decoded_data)):
            sheet[cell + str(i + 1)] = self.decoded_data[i][0]

        wb.save('data/' + self.file_name + ".xlsx")


fin = Wave('IU', 'rb')
fin.show_info()
print()
# fin.show_data()
fin.make_exel()
fin.close()

# fout.close()