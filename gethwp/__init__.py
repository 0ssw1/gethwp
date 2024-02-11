from olefile import OleFileIO
import zlib
import struct
import io
import zipfile
import xml.etree.ElementTree as ET




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


def read_hwpx(hwpx_file_path):
    # 메모리 상에서 ZIP 파일 읽기
    with open(hwpx_file_path, 'rb') as f:
        hwpx_file_bytes = f.read()

    with io.BytesIO(hwpx_file_bytes) as bytes_io:
        with zipfile.ZipFile(bytes_io, 'r') as zip_ref:
            text_parts = []
            for file_info in zip_ref.infolist():
                if file_info.filename.startswith('Contents/') and file_info.filename.endswith('.xml'):
                    with zip_ref.open(file_info) as file:
                        # XML 파일 읽고 파싱하기
                        tree = ET.parse(file)
                        root = tree.getroot()
                        
                        # 모든 텍스트 노드 추출
                        for elem in root.iter():
                            if elem.text:
                                text_parts.append(elem.text.strip())
    
   
    return '\n'.join(text_parts)

def change_word(hwpx_file_path, output_path, find_text, replace_text):
    # 메모리 상에서 수정된 ZIP 파일을 준비하기 위해 BytesIO 객체 생성
    output_bytes_io = io.BytesIO()

    with open(hwpx_file_path, 'rb') as f:
        hwpx_file_bytes = f.read()

    with io.BytesIO(hwpx_file_bytes) as bytes_io:
        with zipfile.ZipFile(bytes_io, 'r') as zip_ref:
            with zipfile.ZipFile(output_bytes_io, 'w') as out_zip:
                for file_info in zip_ref.infolist():
                    with zip_ref.open(file_info) as file:
                        file_contents = file.read()
                        if file_info.filename.startswith('Contents/') and file_info.filename.endswith('.xml'):
                            # XML 파일 처리
                            tree = ET.ElementTree(ET.fromstring(file_contents))
                            root = tree.getroot()
                            
                            # 모든 텍스트 노드 추출 및 수정
                            for elem in root.iter():
                                if elem.text:
                                    elem.text = elem.text.replace(find_text, replace_text)
                            
                            # 수정된 XML 파일을 메모리에 저장
                            modified_file_io = io.BytesIO()
                            tree.write(modified_file_io, encoding='utf-16', xml_declaration=True)
                            modified_file_io.seek(0)
                            file_contents = modified_file_io.read()
                        elif file_info.filename == 'Preview/PrvText.txt':
                            # PrvText.txt 파일 내용 수정
                            modified_text = file_contents.decode('utf-16').replace(find_text, replace_text)
                            file_contents = modified_text.encode('utf-16')

                        # 수정된 파일을 새 ZIP 파일에 추가
                        out_zip.writestr(file_info, file_contents)

    # 수정된 ZIP 파일을 디스크에 저장
    with open(output_path, 'wb') as f:
        f.write(output_bytes_io.getvalue())


if __name__ == "__main__":
    change_word('/Users/0ssw1/Documents/영남일보3.hwpx', '/Users/0ssw1/Documents/영남일보3_modified.hwpx', '인공지능', '인고옹지이느응')