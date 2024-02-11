# Description
HWP 파일 내 텍스트를 문자열 형태로 추출해주는 파이썬 라이브러리입니다.
+ HWPX 파일 치환 기능 추가

Developed by Suh Seungwan(Yumeta lab)
(Thanks to Faith6)

## 사용법
### HWP 파일
```python
import gethwp
hwp = gethwp.read_hwp('test.hwp')
print(hwp)
```

### HWPX 파일
```python
import gethwp
hwp = gethwp.read_hwpx('test.hwpx')
print(hwp)
```

### HWPX 파일 내 텍스트 치환
**HWPX 파일만 가능합니다.**
```python
import gethwp
hwp = gethwp.change_word('test.hwpx','test_output.hwpx','Find Text','Replace Text')
｀｀｀
