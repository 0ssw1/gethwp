from olefile import OleFileIO
import zlib
import struct

def read_hwp(hwp_path):
    with OleFileIO(hwp_path) as ole:
        validate_hwp_file(ole)
        compression_flag = get_compression_flag(ole)
        section_texts = [read_section(ole, section_id, compression_flag) for section_id in get_section_ids(ole)]
    return '\n'.join(section_texts).strip()

def validate_hwp_file(ole):
    required_streams = {"FileHeader", "\x05HwpSummaryInformation"}
    if not required_streams.issubset(set('/'.join(stream) for stream in ole.listdir())):
        raise ValueError("The file is not a valid HWP document.")

def get_compression_flag(ole):
    with ole.openstream("FileHeader") as header_stream:
        return bool(header_stream.read(37)[36] & 1)

def get_section_ids(ole):
    return sorted(
        int(stream[1].replace("Section", "")) 
        for stream in ole.listdir() if stream[0] == "BodyText"
    )

def read_section(ole, section_id, is_compressed):
    with ole.openstream(f"BodyText/Section{section_id}") as section_stream:
        data = section_stream.read()
        if is_compressed:
            data = zlib.decompress(data, -15)
        return extract_text(data)

def extract_text(data):
    text, cursor = "", 0
    while cursor < len(data):
        header = struct.unpack_from("<I", data, cursor)[0]
        type, length = header & 0x3ff, (header >> 20) & 0xfff
        if type == 67:
            text += data[cursor + 4:cursor + 4 + length].decode('utf-16') + "\n"
        cursor += 4 + length
    return text