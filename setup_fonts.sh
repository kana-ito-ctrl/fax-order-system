#!/bin/bash
# Render用: 日本語フォントのインストール
mkdir -p /tmp/fonts
cd /tmp/fonts
# IPAゴシックフォントをダウンロード
if [ ! -f /usr/share/fonts/opentype/ipafont-gothic/ipag.ttf ]; then
    apt-get update -qq && apt-get install -y -qq fonts-ipafont-gothic 2>/dev/null || {
        echo "apt not available, downloading font directly..."
        curl -sL "https://moji.or.jp/wp-content/ipafont/IPAexfont/IPAexfont00401.zip" -o ipa.zip
        unzip -q ipa.zip
        mkdir -p ~/.fonts
        cp IPAexfont00401/*.ttf ~/.fonts/
        fc-cache -f ~/.fonts 2>/dev/null || true
    }
fi
echo "Font setup complete"
