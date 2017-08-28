# -*- coding: UTF-8 -*-
###文本格式转换器
#针对多格式文本。首先全部转换为纯文本，再对字符进行基本去噪。
#可处理格式包括：pdf（pdfminer），docx（docxpy），html，txt（尚不能处理格式：mht，htm，doc）

import sys, getopt
import re
import os
import docxpy

opts, args = getopt.getopt(sys.argv[1:], "ahi:o:", ["all","help", "input=", "output="])

def usage():
   print 'python '+ sys.argv[0] + ' -i input_file -o output_file'
   print 'python '+ sys.argv[0] + ' -a'

def strdecode(sentence):
    if not isinstance(sentence, unicode):
        try:
            sentence = sentence.decode('utf-8')
        except UnicodeDecodeError:
            sentence = sentence.decode('gbk', 'ignore')
    return sentence

#字符预处理规则
def char_preprocess(txt):
   txt=strdecode(txt)
   symbol_en='[~|`|!|@|#|$|%|^|&|*|(|)|_|-|+|=|\{|\[|}|\]|\||\\\|:|;|"|\'|<|,|>|.|?|/| ]'
   symbol_ch=strdecode('[：|；|“|”|’|‘|《|》|，|。|？|【|】|）|（]')
   single_ch=strdecode('︳')
   #换行含义字符归一化
   txt=re.sub('[\t|\r|\|]','\n',txt)
   txt=re.sub(single_ch,'\n',txt)
   txt=re.sub('(?<=[^[ |\n]]) {3,}(?=[^ ])','\n',txt)
   #删除空格行+符号行
   txt=re.sub('(?<=\n)'+symbol_en+'+(?=\n)','',txt)
   txt=re.sub('(?<=\n)'+symbol_ch+'+(?=\n)','',txt)
   #删除空行
   txt=re.sub('\n(?=\n)|$\n','',txt)
   return txt.encode('u8')

def pdf(i,o,ip,op):
   print 'pdf'
   os.system('python pdf2txt.py %s > %s'%(ip+i,op+o))
   f=open(op+o,'r')
   content=f.read()
   f.close()
   content_filter=char_preprocess(content)
   f=open(op+o,'w')
   f.write(content_filter)
   f.close()

def docx(i,o,ip,op):
   print 'docx'
   text = docxpy.process(ip+i)
   f=open(op+'%s'%o,'w')
   text=char_preprocess(text)
   f.write(text)
   f.close()

def htm(i,o,ip,op):
   print 'html'
   f=open(ip+i,'r')
   content=f.read()
   content_filter=re.sub('<.*?>','',content)
   f.close()
   f=open(op+'%s'%o,'w')
   content_filter=char_preprocess(content_filter)
   #针对htm格式的特殊字符处理
   content_filter=re.sub('&nbsp','',content_filter)
   ###
   f.write(content_filter)
   f.close()

def txt(i,o,ip,op):
   print 'txt'
   f=open(ip+'%s'%i,'r')
   content=f.read()
   f.close()
   content_filter=char_preprocess(content)
   f=open(op+'%s'%o,'w')
   f.write(content_filter)
   f.close()

def run(input_file,output_file,input_path,output_path):
   file_list=re.split('\.',input_file)
   if len(file_list)<>2:
      print 'File name invalid'
      exit()
   else:
      if file_list[1]=='pdf':
         pdf(input_file,output_file,input_path,output_path)
      # elif file_list[1]=='doc':
      #    subprocess.call(['soffice', '--headless', '--convert-to', 'docx', input_path+input_file])
      #    docx(input_file,output_file,input_path,output_path)
      elif file_list[1]=='docx':
         docx(input_file,output_file,input_path,output_path)
      elif file_list[1]=='html':
         htm(input_file,output_file,input_path,output_path)
      elif file_list[1]=='txt':
         txt(input_file,output_file,input_path,output_path)
      else:
         print 'File format invalid'
         exit()

input_file, output_file = '', ''
input_path,output_path='',''
for op, value in opts:
   if op == '-i':
          input_file = value
          print input_file
   elif op == '-o':
          output_file = value
          print output_file
   elif op == '-h':
          usage()
          sys.exit()
   elif op == '-a':
          input_path='format_free_input/'
          output_path='format_free_output/'
          output_count=0
          for input_file in os.listdir(input_path):
             output_count+=1
             output_file='%d'%output_count
             run(input_file,output_file,input_path,output_path)
          sys.exit()

run(input_file,output_file,input_path,output_path)