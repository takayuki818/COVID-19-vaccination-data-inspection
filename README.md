# Name
新型コロナウイルスワクチン接種データの適否検査（R5.11.1時点）
## Overview
新型コロナウイルスワクチン接種データを「使用ワクチン」「接種日」「年齢・回数・接種間隔」別に分岐検査し、適否を返します。
## Note
・あくまでワクチン単位での使用適否検査であり、法令等の全てを網羅した検査ではありません。  
・乳幼児の「生後6か月未満」判定には対応していません。
・令和5年11月1日時点の制度内容までを反映
## Features
引数として「年齢, 接種日, ワクチン名, 回数, 前回接種日, 前回年齢エラー文」を投入することで、  
「年齢エラー文」「接種（回数・間隔）エラー文」を返します。  
※「前回年齢エラー文」は、乳幼児・小児接種における  
　「1回目接種後の加齢により制限年齢を超えた場合」にOKエラーと判別する必要がある場合のみ参照しています。
