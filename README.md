# VBAPainter
Make Excel grid sheet and read RGB files to plot colors on Excel

# How to use
.basファイルをエクセルファイルにインポートします。
画像のRGB値それぞれの値をマッピングした、0~255の値が並んだcsvファイル(カンマ区切り)を用意します。
マクロ"importRGBTextToBGColor"を実行すると、R, G, Bの順でcsvファイルを選択する画面が出ます。
正しい順番でファイルを選択すると、エクセル上に画像が表示されます。
順番を間違えると、RGBのいずれかが反転した画像が表示されます。

マクロ"adjustCellSize"を実行すると、セルサイズをエクセル方眼紙サイズにします。.basファイル内の定数pxを変えることで、方眼紙のサイズを変更できます。
