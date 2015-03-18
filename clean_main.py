# -*- coding: utf8 -*-
__author__ = 'SK_Lee'
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import re

import xlrd

# 中文
data = xlrd.open_workbook('Survey_raw_data_2015011211304403825.xls')

# 英文
# data = xlrd.open_workbook('Survey_raw_data_2015011211415963032.xls')
table = data.sheet_by_name('raw data')


import xlsxwriter

outxlsx = 'clean_cn_data.xlsx'
workbook = xlsxwriter.Workbook(outxlsx)
worksheet = workbook.add_worksheet('clean data')



def q1(qcol,move,worksheet):
    # 寫下問題
    worksheet.write_column(0,move,qcol)
    qdict={}
    # 分開數字跟問題
    for irow,col in enumerate(qcol[1:]):
        col=col.strip()
        if col:
            if qdict.get(col,'empty')=='empty':
                qdict[col]=[irow+1]
            else:
                qdict[col].append(irow+1)
    # 排序問題
    sorted_q=sorted(qdict.iteritems(), key=lambda key_value: int(key_value[0].split(')')[0]))
    # sorted_q = sorted(qdict.items(), key=operator.itemgetter(0))
    for iq in sorted_q:
        move+=1
        worksheet.write_string(0,move,iq[0])
        for v in iq[1]:
            worksheet.write_string(v,move,'v')
    return move


def qm(qcol,move,worksheet):
    worksheet.write_column(0,move,qcol)
    qdict={}
    # 分開數字跟問題
    for irow,col in enumerate(qcol[1:]):
        col=col.strip()
        mcol=col.split(';')
        for m in mcol:
            if m:
                m=m.strip('\t')
                if qdict.get(m,'empty')=='empty':
                    qdict[m]=[irow+1]
                else:
                    qdict[m].append(irow+1)
    # 排序問題
    sorted_q=sorted(qdict.iteritems(), key=lambda key_value: int(key_value[0].split(')')[0]))
    # sorted_q = sorted(qdict.items(), key=operator.itemgetter(0))
    for iq in sorted_q:
        move+=1
        worksheet.write_string(0,move,iq[0])
        for v in iq[1]:
            worksheet.write_string(v,move,'v')
    return move


def qms(qcol,move,worksheet):
    worksheet.write_column(0,move,qcol)
    move+=1
    title=[]
    qdict={}
    has_title=0
    for irow,col in enumerate(qcol[1:]):
        irow+=1
        col=col.strip()
        if col:
            term=col.split(';')
            for iterm in term:
                if not has_title:
                    title.append(iterm.split(':')[0].strip())
                val=re.findall('\[(\d+)\]', iterm)
                if qdict.get(irow,'empty')=='empty':
                    qdict[irow]=val
                else:
                    qdict[irow].extend(val)
            has_title=1
        else:
            qdict[irow]=[-1]

    for irow in xrange(table.nrows):
        for icol in xrange(len(title)):
            if irow==0:
                worksheet.write_string(irow,move+icol,title[icol])
            elif qdict[irow][0]!=-1:
                worksheet.write_number(irow,move+icol,int(qdict[irow][icol]))
    move+=(len(title)-1)
    return move



def mapcol(qcol,move,worksheet,rule):
    rule='Country:TAIWAN,R.O.C.;Region:TWN;language:tw'
    rules=rule.split(';')
    for srule in rules:
        worksheet.write_string(0,move,srule.split(':')[0])
        row=1
        for irow in qcol[1:]:
            if irow:
                worksheet.write_string(row,move,srule.split(':')[1])
                row+=1
        move+=1
    move-=1
    return move


icol=0
move=0
not_clean=['EMail','Name','Address','Zip Code','Tel','Country','Filling Data','Case NO','IP','建議','suggestion']


while icol< table.ncols:
    qname=table.cell_value(0,icol)
    colval=table.col_values(icol)

    noclean=False
    for no in not_clean:
        if no in qname:
            noclean=True
            break

    # if qname in ['Country','年齡']:
    #     move=mapcol(colval,move,worksheet,"")
    if noclean:
        worksheet.write_column(0, move, colval)
    else:
        if 'Please score each item' in qname or '本題每項都需要勾選一個分數' in qname:
            move=qms(colval,move,worksheet)
        elif 'Select all that apply' in qname or '複選' in qname:
            move=qm(colval,move,worksheet)
        else:
            move=q1(colval,move,worksheet)

    icol+=1
    move+=1

workbook.close()

