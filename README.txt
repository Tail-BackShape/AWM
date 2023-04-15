AWM_GUI.py
(Auto Wing Maker GUI version)

Inventorなどのソフトに読み込ませるExcelの翼型データを自動生成するプログラム

使用前にPowerShellやコマンドプロンプトなどのターミナルで
pip install openpyxl
と打ち、openpyxlを使用できるようにする。

CUI versionと違い、元翼データを変えたいときにいちいちプログラムを書き直す必要がない。

Openで元翼データ(CSV形式*)を選択し、wing spanで翼弦長を設定する。
ファイル保存はSaveを押して任意のファイルに保存することができる。

*CSV形式
・A列:x
・B列:y

下記のようなサイトを参考に元翼データを作る。
http://airfoiltools.com/
