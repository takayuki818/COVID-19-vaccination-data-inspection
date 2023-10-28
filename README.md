# COVID-19-vaccination-data-inspection
新型コロナウイルスワクチン接種データの適否検査（R5.9.25時点）  
※あくまでワクチン単位での使用適否検査であり、法令等の全てを網羅した検査ではありません。  
※乳幼児の「生後6か月未満」判定には対応していません。
## 機能概要
引数として「年齢, 接種日, ワクチン名, 回数, 前回接種日, 前回年齢エラー文」を投入することで、  
「年齢エラー文」「接種（間隔）エラー文」を返します。  
※「前回年齢エラー文」は乳幼児・小児接種において、1回目接種後の加齢により年齢判定範囲を超えた場合にOKエラーとするために参照。
## 分岐処理
ワクチン名分岐  
　→　接種日分岐（制度改正時期参照）  
　　→　年齢適否、接種（回数・間隔）適否を判定  
