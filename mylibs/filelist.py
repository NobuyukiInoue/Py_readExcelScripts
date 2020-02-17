# -*- coding: utf-8 -*-

import os

def find_all_files(directory, Pattern, Exclude):
    res = []
    words_match = Pattern.split("*")
    words_exclude = Exclude.split("*")
    for root, dirs, files in os.walk(directory):
        isExclude = False

        # Excludeのパターンを含むディレクトリの場合は次へ
        for word in words_exclude:
            if word == "":
                continue
            if word in root:
                isExclude = True
                break

        if isExclude:
            continue

        # ファイル名にPatternを含むかどうか調べる
        for file in files:
            isMatched = False
            for word in words_match:
                if word == "":
                    continue
                if word in file:
                    isMatched = True
                    continue
                else:
                    isMatched = False
                    break
            if isMatched:
                res.append(os.path.join(root, file))

    return res
