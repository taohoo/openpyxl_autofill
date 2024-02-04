# -*- coding: utf-8 -*-
from setuptools import setup, find_packages

with open('README.MD', 'r', encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='openpyxl_autofill',
    version='0.2.2',
    description='针对openpyxl的功能扩展',
    long_description=long_description,
    long_description_content_type='text/markdown',
    author='HuBo',
    author_email='taohoo@163.com',
    url='https://github.com/taohoo/openpyxl_autofill',
    license='MIT License',
    packages=find_packages(),
    platforms=['all'],
    python_requires='>=3.8',
    include_package_data=True,  # 打包python发行包中的include和libs，对应配置在MANIFEST.in中，
    install_requires=['openpyxl'],    # 最简单'certifi'，也可以这样写'urllib3>=1.21.1,<1.27'
    classifiers=[
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: Implementation :: CPython',
        'Programming Language :: Python :: Implementation :: PyPy'
    ]
)
