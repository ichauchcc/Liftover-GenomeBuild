#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed May 12 09:43:21 2021

@author: yuchen
"""

# install the package 'pyliftover'
# pip install pyliftover
# install docx
# pip install python_docx

import docx
from pyliftover import LiftOver

# Change the path to where you store the file
document = docx.Document('/Users/yuchen/Desktop/DAG1.docx')

# Change the version you would like to lift over
# current is hg19 to hg38 
lo = LiftOver('hg19', 'hg38')
for paragraph in document.paragraphs:
    i=0
    for run in paragraph.runs:
        for seg in paragraph.text.split(' '):
            if i>0:
                i-=1
                continue
            if seg.isdigit():
                # Remember to change the chrosome to the right one
                a=lo.convert_coordinate('chr3',int(seg))
                if not a:
                    print('Wrong input: {}'.format(seg))
                    i=2
                    continue
                newseg=str(a[0][1])
                run.text = run.text.replace(seg,newseg)

# Change the path you store the file and give it a name
document.save('/Users/yuchen/Desktop/DAG1_hg38.docx')
