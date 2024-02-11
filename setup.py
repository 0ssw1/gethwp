from setuptools import setup, find_packages

with open('README.md', encoding='utf-8') as f:
    long_description = f.read()

setup(name='gethwp', # 패키지 명

version='1.1.1',

long_description= long_description,
long_description_content_type='text/markdown',

description='Get text from HWP/HWPX file',

author='Suh Seungwan',

author_email='seo@seowan.net',

url='https://github.com/0ssw1/gethwp',

license='MIT', 

python_requires='>=3',

install_requires=['olefile'], 

packages=['gethwp'] 

)