# 中文 .docx 文件查重工具
本工具采用随机算法计算指定文件夹内两两 .docx 文件间的相似性。算法原理详见 release 页面的 .pdf 文件（from NYU）。

## 运行依赖安装
```
pip install python-docx thulac progressbar numpy
```

## 程序运行
```
usage: check_docx_similarity.py [-h] --dir DIR --out OUT
                                [--hash-width HASH_WIDTH]
                                [--hash-step HASH_STEP]
                                [--sample-cnt SAMPLE_CNT]
```

## 参数说明
#### `--dir`
指定包含 .docx 文件的文件夹
#### `--out`
指定输出 .csv 文件的路径
#### `--hash-width`
指定每个哈希块的词长，默认连续8个中文词语构成一个哈希块。增大该值会使 false positive 概率变小，但 false negative 概率变大。该参数几乎不影响计算速度。
#### `--hash-step`
指定相邻两个哈希块之间的间距，默认直接相邻（步长1）。增大该值会使 false positive 概率变小，但 false negative 概率变大。增大该参数能加快计算速度。
#### `--sample-cnt`
随机采样数。对于每个 .docx 文件，本程序生成一个高维向量，通过该高维向量间的相似性来判断文件相似性。向量的每一维均通过随机采样生成。此参数控制了高维向量的维数。增大该值会使结果更可靠，但显著增加计算时间。
