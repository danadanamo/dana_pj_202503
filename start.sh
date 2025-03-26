#!/bin/bash

# 仮想環境が存在しない場合は作成
if [ ! -d "venv" ]; then
    echo "仮想環境を作成中..."
    python3 -m venv venv
fi

# 仮想環境をアクティブにする
source venv/bin/activate

# 依存関係をインストール
echo "依存関係をインストール中..."
pip install -r requirements.txt

# アプリケーションを起動
echo "アプリケーションを起動中..."
python step_02/d_pj_image_grid_app_v0.03.py 