a = Analysis(['app.py'], pathex=['.'], hiddenimports=['tkinter'], noarchive=False)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name='Общий журнал', console=False, upx=False)
