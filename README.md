# rpy2xlsx
From "/tl/language/*.rpy" to "output/*.xlsx"(like T++)
把renpy SDK生成的 "/game/tl/language" 内的 "*.rpy" 转化为类似T++的excel形式，以使用其他工具进行翻译。

使用方法：
1.输入将文件夹放入./rpy_files，运行r2e.py，在./tran_files/EXCEL_FILE转化为excel文件(用作AiNiee翻译)，在./tran_files/JSON_FILE储存行数据json文件
2.进行翻译，如AiNiee
3.将翻译后的excel目录覆盖至./tran_files/EXCEL_FILE，运行e2r，结果会写入回./rpy_files
4../rpy_files目录内的language文件夹覆盖游戏文件/game/tl/language

PS:由于T++导入导出总是卡死，大怒，遂尝试自己编一个，第一次编程，目前代码变量乱七八糟，不过好在能用
