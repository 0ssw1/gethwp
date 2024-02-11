from setuptools import setup, find_packages

setup(name='gethwp', # 패키지 명

version='1.0.0',

description='Get text from HWP file',

author='Suh Seungwan',

author_email='seo@seowan.net',

url='http://seowan.net',

license='MIT', 

python_requires='>=3',

install_requires=['olefile'], 

packages=['gethwp'] 

)