#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
import sys
from com.sun.star.beans import UnknownPropertyException
def macro():
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	doc = XSCRIPTCONTEXT.getDocument()
	if not doc.supportsService("com.sun.star.text.TextDocument"):  # Writerドキュメントに結果を出力するのでWriterドキュメントであることを確認する。
		raise RuntimeError("Please execute this macro with a Writer document.")  # Writerドキュメントでないときは終わる。	
	txts = []  # Writerドキュメントに出力する文字列を入れるリスト。
	predefinedpathproperties = "Addin",\
								"AutoCorrect",\
								"AutoText",\
								"Backup",\
								"Basic",\
								"Bitmap",\
								"Config",\
								"Dictionary",\
								"Favorite",\
								"Filter",\
								"Gallery",\
								"Graphic",\
								"Help",\
								"Linguistic",\
								"Module",\
								"Palette",\
								"Plugin",\
								"Storage",\
								"Temp",\
								"Template",\
								"UIConfig",\
								"UserConfig",\
								"UserDictionary",\
								"Work"  # https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Path_Settings
	pathsettingsservice = ctx.getByName('/singletons/com.sun.star.util.thePathSettings')
	for predefinedpathproperty in predefinedpathproperties:
		try:
			url = pathsettingsservice.getPropertyValue(predefinedpathproperty)
		except UnknownPropertyException:
			t = "Cannot get a property value of {}.".format(predefinedpathproperty)
			print(t, file=sys.stderr)
			txts.append(t)
			continue
		systempath = ";".join([os.path.normpath(unohelper.fileUrlToSystemPath(p)) if p.startswith("file://") else p for p in url.split(";")])  # 各パスについてシステムパスにして正規化する。
		t = "{} : {}".format(predefinedpathproperty, systempath)
		print(t)
		txts.append(t)
	for predefinedpathproperty in predefinedpathproperties:  # プロパティ名に_writableをつけてみる。パスが取得できたもののみ出力。
		predefinedpathproperty = "{}_writable".format(predefinedpathproperty)
		try:
			url = pathsettingsservice.getPropertyValue(predefinedpathproperty)
		except:
			continue
		if url:		
			systempath = ";".join([os.path.normpath(unohelper.fileUrlToSystemPath(p)) if p.startswith("file://") else p for p in url.split(";")])  # 各パスについてシステムパスにして正規化する。
			t = "{} : {}".format(predefinedpathproperty, systempath)
			print(t)
			txts.append(t)
	t = "These paths have been converted to system paths."
	print(t)
	txts.append(t)
	doc.getText().setString("\n".join(txts))  # Writerドキュメントに出力。 				
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
	import officehelper
	from functools import wraps
# 	import sys
	from com.sun.star.beans import PropertyValue
	from com.sun.star.script.provider import XScriptContext  
	def connectOffice(func):  # funcの前後でOffice接続の処理
		@wraps(func)
		def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
			try:
				ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
			except:
				print("Could not establish a connection with a running office.", file=sys.stderr)
				sys.exit()
			print("Connected to a running office ...")
			smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
			print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
			return func(ctx, smgr)  # 引数の関数の実行。
		def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
			cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
			node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
			ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
			return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
		return wrapper
	@connectOffice  # mainの引数にctxとsmgrを渡すデコレータ。
	def main(ctx, smgr):  # XSCRIPTCONTEXTを生成。
		class ScriptContext(unohelper.Base, XScriptContext):
			def __init__(self, ctx):
				self.ctx = ctx
			def getComponentContext(self):
				return self.ctx
			def getDesktop(self):
				return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
			def getDocument(self):
				return self.getDesktop().getCurrentComponent()
		return ScriptContext(ctx)  
	XSCRIPTCONTEXT = main()  # XSCRIPTCONTEXTを取得。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントを取得。
	if not hasattr(doc, "getCurrentController"):  # ドキュメント以外のとき。スタート画面も除外。
		XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/swriter", "_blank", 0, ())  # Writerのドキュメントを開く。
		while doc is None:  # ドキュメントのロード待ち。
			doc = XSCRIPTCONTEXT.getDocument()
	macro()