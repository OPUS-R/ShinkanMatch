import subprocess
import os
# pip install pyinstaller

#exe化するpyファイルのパス
target_file = "convert.py"

#コマンド(オプション)
command = [
    "pyinstaller",
    "--onefile",       #exeを1つにする
    "--noconsole",     #コンソールを表示しない(GUIを表示する為)
    target_file
]


print(f"Running: {' '.join(command)}")
subprocess.run(command)#実行

#distフォルダの下に生成
exe_path = os.path.join("dist", os.path.splitext(os.path.basename(target_file))[0] + ".exe")
if os.path.exists(exe_path):
    print(f"exe化成功: {exe_path}")
else:
    print("❌ 失敗")
