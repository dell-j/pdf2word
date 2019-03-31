import PyPDF2
import csv
import sys,os,re,collections,nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem.wordnet import WordNetLemmatizer
import pdf2txt

# 正则表达式过滤特殊符号用空格符占位，双引号、单引号、句点、逗号
pat_letter = re.compile(r'[^a-zA-Z \']+')
# 还原常见缩写单词
pat_is = re.compile("(it|he|she|that|this|there|here)(\'s)", re.I)
pat_s = re.compile("(?<=[a-zA-Z])\'s") # 找出字母后面的字母
pat_s2 = re.compile("(?<=s)\'s?")
pat_not = re.compile("(?<=[a-zA-Z])n\'t") # not的缩写
pat_would = re.compile("(?<=[a-zA-Z])\'d") # would的缩写
pat_will = re.compile("(?<=[a-zA-Z])\'ll") # will的缩写
pat_am = re.compile("(?<=[I|i])\'m") # am的缩写
pat_are = re.compile("(?<=[a-zA-Z])\'re") # are的缩写
pat_ve = re.compile("(?<=[a-zA-Z])\'ve") # have的缩写

def replace_abbreviations(text):
    new_text = pat_letter.sub(' ', text).strip().lower()
    new_text = pat_is.sub(r"\1 is", new_text)
    new_text = pat_s.sub("", new_text)
    new_text = pat_s2.sub("", new_text)
    new_text = pat_not.sub(" not", new_text)
    new_text = pat_would.sub(" would", new_text)
    new_text = pat_will.sub(" will", new_text)
    new_text = pat_am.sub(" am", new_text)
    new_text = pat_are.sub(" are", new_text)
    new_text = pat_ve.sub(" have", new_text)
    new_text = new_text.replace('\'', ' ')
    return new_text

#文本分词，并根据规则过滤掉标点符号及连字符等
def participle(text):
    text = replace_abbreviations(text)
    #使用word_tokenize 需要下载 nltk.download('punkt')
    #保存在 nltk_data\tokenizers\punkt
    tokens = word_tokenize(text)
    punctuations = ['a','b','c','d','e','f','g','h','i','j','k','l','m','o','p','q','r','s','t','u','v','w','x','y','z']
    #使用stopwords 需要下载 nltk.download('stopwords')
    #保存在 nltk_data\corpora\stopwords
    stop_words = stopwords.words('english')
    keywords = [word for word in tokens if (word not in stop_words) and (word not in punctuations)]
    return keywords

# pos和tag有相似的地方，通过tag获得pos
def get_wordnet_pos(treebank_tag):
    if treebank_tag.startswith('J'):
        return nltk.corpus.wordnet.ADJ
    elif treebank_tag.startswith('V'):
        return nltk.corpus.wordnet.VERB
    elif treebank_tag.startswith('N'):
        return nltk.corpus.wordnet.NOUN
    elif treebank_tag.startswith('R'):
        return nltk.corpus.wordnet.ADV
    else:
        return ''

# 标记词性，然后将词形还原
def tagging(keywords):
    lmtzr = WordNetLemmatizer()
    #使用词性标注功能，需下载 nltk.download('averaged_perceptron_tagger')
    #保存在nltk_data\taggers\averaged_perceptron_tagger
    #还需下载 nltk.download('wordnet')
    #保存在 nltk_data\corpora\wordnet
    tags = nltk.pos_tag(keywords)
    new_words = []
    for tag in tags:
        pos = get_wordnet_pos(tag[1])
        if pos:
            lemmatized_word = lmtzr.lemmatize(tag[0], pos)
            new_words.append(lemmatized_word)
        else:
            new_words.append(tag[0])
    return new_words

#按词频统计排序，返加元组列表
def word_freq(new_words):
    c_words = collections.Counter(new_words)
    c_list = c_words.most_common() #most_common(N)将counter字典转为元组列表，取前N个
    new_words = []
    for item in c_list:
        new_words.append(item[0])
    return new_words

#读取词汇表
def read_csv(filename):
    dicts = []
    with open(filename,'r') as csvfile:
        rows = csv.reader(csvfile, delimiter = ',')
        for row in rows:
            dicts.append(row[0])
    return dicts

#按字频查找词汇表（页面字频高的在前）
def lookup_dict(dicts,new_words):
    word_m = []
    for word in new_words:
        try:
            _ = dicts.index(word)
            word_m.append(word)
        except:
            _ = 0
    return word_m

#读取PDF文件提取文本
def read_pdf(filename,start,end):
    t_list = []
    if os.path.exists(filename):
        #(filepath,tempfilename) = os.path.split(filename)
        #(shotname,extension) = os.path.splitext(tempfilename)
        (_,tempfilename) = os.path.split(filename)
        (_,extension) = os.path.splitext(tempfilename)
    else:
        extension = ''
    if extension == '.pdf':
        try:
            with open(filename,'rb') as pdfFileObj:
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                total_pages = pdfReader.numPages
            if (start-1) >= 0 and (end-1) <= total_pages and end >= start:
                count = end-start+1
                page_numbers = [start-1+i for i in range(count)]
                #从pdf2txt模块获取返回文件句柄和文本内容列表
                outfp,t_list = pdf2txt.extract_text(files=[filename], page_numbers=page_numbers)
                outfp.close()
                #for i in range(count):
                    #pageObj = pdfReader.getPage(start-1+i)
                    #t_list.append(pageObj.extractText())
            elif (end-1) > total_pages:
                print("页码超出最大范围")
            elif (start-1) < 0:
                print("起始页码不正确")
            elif end < start:
                print("终止页码小于起始页码")
        except OSError as err:
            print(err)
        except:
            print("文件读写未知错误！")
    elif extension == '.txt':
        print(".txt 文件，尚未支持！")
    elif extension == '':
        print(filename+"文件不存在！")
    else:
        print("程序只能处理 pdf文件！")
    return t_list

#保存文本文件
def save_text(filename,text):
    with open(filename,'w') as text_file:
        text_file.write(text)

def pdf_to_word(filename,page_start=1,page_end=1):
    #filename = 'calculus.pdf'
    #page_start = 1 #以1为起点
    #page_end = 3
    (_,tempfilename) = os.path.split(filename)
    (shotname,_) = os.path.splitext(tempfilename)
    t_list = read_pdf(filename,page_start,page_end)
    count = 0
    total_text = ""
    if t_list != []:
        for text in t_list:
            if text.strip() != '':
                keywords = participle(text)
                new_words = tagging(keywords)
                new_words = word_freq(new_words)
                list_m = lookup_dict(read_csv('middle.csv'),new_words)   #高中词汇
                list_c4 = lookup_dict(read_csv('cet4.csv'),new_words)
                list_c4 = [word for word in list_c4 if word not in list_m]  #四级词汇（不包括高中词汇）
                list_c6 = lookup_dict(read_csv('cet6.csv'),new_words)
                list_c6 = [word for word in list_c6 if (word not in list_m) and (word not in list_c4)] #六级词汇（不包括四、高中）
                list_g = lookup_dict(read_csv('gre.csv'),new_words)
                list_g = [word for word in list_g if (word not in list_m) and (word not in list_c4) and (word not in list_c6)] #GRE词汇（不包刮四六高）
                list_ox = lookup_dict(read_csv('oxford.csv'),new_words)
                list_ox = [word for word in list_ox if (word not in list_m) and (word not in list_c4) and (word not in list_c6) and (word not in list_g)] #用简明牛津词典库垫底
                title = '#'+shotname+'_Page'+str(page_start+count)+'\n'
                total_text = total_text + title + ('\n').join(list_m + list_c4 +list_c6 + list_g + list_ox) + '\n'
            count = count + 1
        save_filename = shotname+'_P'+str(page_start)+'-P'+str(page_end)+'.txt'
        save_text(save_filename,total_text)
    else:
        print("内容为空或取不到页面文本！")
    return

def main(argv):
    if len(argv) < 4:
        print('请输入读取文件名和页号起止，例：pdf_to_word.py calculus.pdf 1 3')
    else:
        try:
            p1 = int(argv[2])
            p2 = int(argv[3])
        except:
            print('输入类型错误，请输入整数类型的页码！')
            return    
        pdf_to_word(argv[1],p1,p2)

if __name__ == '__main__':
    main(sys.argv)

