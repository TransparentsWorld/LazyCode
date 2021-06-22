import jieba
import jieba.posseg as pseg #词性标注
import jieba.analyse as anls #关键词提取

jieba.load_userdict(r"C:\Users\m\Desktop\jieba分词字典.txt")

filename = r"C:\Users\m\Desktop\日记.txt"
with open(filename,"r",encoding="utf-8") as f:
    diary_list = f.read()

seg_list = jieba.cut_for_search(diary_list,HMM=True)

data1 = {}
for chara in seg_list:
    if len(chara) < 2:
        continue
    if chara in data1:
        data1[chara] += 1
    else:
        data1[chara] = 1

data1 = sorted(data1.items(), key=lambda x: x[1], reverse=True)  # 排序

for data in data1:
    print(data)