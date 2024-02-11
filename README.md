# Description
HWP 파일 내 텍스트를 문자열 형태로 추출해주는 파이썬 라이브러리입니다. 유메타랩(주) 대표 서승완이 개발하였으며, 무료(MIT 라이선스)로 배포됩니다. (Thank to Faith6)

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

