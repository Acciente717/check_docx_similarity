import os
import sys
import docx
import thulac
import hashlib
import argparse
import progressbar
import numpy as np


def init_progress_bar(maxval):
    global BAR
    BAR = progressbar.ProgressBar(maxval=maxval,
                                  widgets=[progressbar.Bar('=', '[', ']'),' ', progressbar.Percentage()])
    print('', end='', flush=True)


def tokens_of_file(path, thu):
    doc = docx.Document(path)
    string = ''.join([par.text.strip() for par in doc.paragraphs])

    tokens = thu.cut(string, text=True)
    
    tokens = [token for token in tokens.split(' ')
              if token not in ('，', '。', '：', '；', '《', '》', '、', ',', '.', ';', ':',
                               '\'', '"', '‘', '’', '“', '”', '(', ')', '（', '）', 
                               '我', '得', '的', '地', '了', '会', '是', '你', '他', '她' ,'它')]
    return tokens


def hash_tokens(tokens, width, step):
    return [
        hashlib.md5(''.join(tokens[idx:idx+width]).encode('utf-8')).digest()
        for idx in range(0, len(tokens)-width, step)
    ]


def number_digests(digests):
    return {
        digest : number
        for digest, number in zip(digests, range(len(digests)))
    }


def minhash_vec_from_matrix(matrix):
    np.random.shuffle(matrix)
    row, col = matrix.shape
    
    vec = np.zeros(col, dtype=np.uint32)
    for c in range(col):
        for r in range(row):
            if matrix[r, c]:
                vec[c] = r
                break
    return vec


def minhash_matrix_from_matrix(matrix, hash_cnt):
    vecs = []
    init_progress_bar(hash_cnt)
    BAR.start()
    for i in range(hash_cnt):
        vecs.append(minhash_vec_from_matrix(matrix))
        BAR.update(i)
    minhash_matrix = np.vstack(vecs)
    BAR.finish()
    return minhash_matrix


def similarity(vec1, vec2):
    return sum(vec1 == vec2)


def pairwise_similarity(minhash_matrix):
    _, col = minhash_matrix.shape
    return [
        (left, right, similarity(minhash_matrix[:,left], minhash_matrix[:,right]))
        for left in range(col-1)
        for right in range(left+1, col)
    ]


def parse_args():
    parser = argparse.ArgumentParser(description='Check pairwise similarities of .docx files.')
    parser.add_argument('--dir', help='the directory containing .docx files', required=True)
    parser.add_argument('--out', help='the output file path', required=True)
    parser.add_argument('--hash-width', help='the word length of a hashing block (default 8)', default=8)
    parser.add_argument('--hash-step', help='the word step between hashing block (default 1)', default=1)
    parser.add_argument('--sample-cnt', help='sample count (default 1000)', default=1000)

    args = parser.parse_args()
    args.thu = thulac.thulac(seg_only=True)
    return args


def main():
    args = parse_args()

    filenames = [file for file in os.listdir(args.dir)
                 if os.path.splitext(file)[1] == '.docx']

    all_digests = []
    filename_to_digests = {}

    print('Digesting %d files...' % len(filenames), flush=True)
    init_progress_bar(len(filenames))
    BAR.start()
    for idx, filename in enumerate(filenames):
        path = os.path.join(args.dir, filename)
        digests = set(hash_tokens(tokens_of_file(path, args.thu), args.hash_width, args.hash_step))
        all_digests.append(digests)
        filename_to_digests[filename] = digests
        BAR.update(idx)
    BAR.finish()
    print('', end='', file=sys.stderr, flush=True)
    print(flush=True)

    all_digests = sorted(list(set.union(*all_digests)))
    digest_to_number = number_digests(all_digests)

    matrix = np.zeros((len(all_digests), len(filenames)), dtype=np.bool)

    print('Calculating MinHashes...', flush=True)
    for idx, filename in enumerate(filenames):
        for digest in filename_to_digests[filename]:
            row_num = digest_to_number[digest]
            matrix[row_num, idx] = True

    minhash_matrix = minhash_matrix_from_matrix(matrix, args.sample_cnt)
    print('', end='', file=sys.stderr, flush=True)
    print(flush=True)

    print('Calculating similarities between files...', flush=True)
    similarities = pairwise_similarity(minhash_matrix)
    similarities.sort(key=lambda x: x[2], reverse=True)

    print('Writing to %s ...' % args.out, flush=True)
    with open(args.out, 'wt') as f:
        print('file1,file2,similarity', file=f)
        for left, right, sim in similarities:
            sim = sim * 100 / args.sample_cnt
            print('%s,%s,%.1f%%' % (filenames[left], filenames[right], sim), file=f)
    print('Done.', flush=True)


if __name__ == '__main__':
    main()
