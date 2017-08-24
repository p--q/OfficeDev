#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
import sys
from com.sun.star.beans import UnknownPropertyException
def main(ctx, smgr):  # ctx: コンポーネントコンテクスト、smgr: サービスマネジャー
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
								"Work"
	pathsettingsservice = ctx.getByName('/singletons/com.sun.star.util.thePathSettings')
	for predefinedpathproperty in predefinedpathproperties:
		try:
			url = pathsettingsservice.getPropertyValue(predefinedpathproperty)
		except UnknownPropertyException:
			print("Cannot get a property value of {}.".format(predefinedpathproperty), file=sys.stderr)
			continue
		systempath = os.path.normpath(unohelper.fileUrlToSystemPath(url)) if url else url  # urlが空欄のときはパスを変換しない。
		print("{} : {}".format(predefinedpathproperty, systempath))
	for predefinedpathproperty in predefinedpathproperties:  # プロパティ名に_writableをつけてみる。パスが取得できたもののみ出力。
		predefinedpathproperty = "{}_writable".format(predefinedpathproperty)
		try:
			url = pathsettingsservice.getPropertyValue(predefinedpathproperty)
		except:
			continue
		if url:		
			systempath = os.path.normpath(unohelper.fileUrlToSystemPath(url))
			print("{} : {}".format(predefinedpathproperty, systempath))
if __name__ == "__main__":
	import officehelper
	from functools import wraps
# 	import sys
	from com.sun.star.beans import PropertyValue
	# funcの前後でOffice接続の処理
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
	main = connectOffice(main)
	main()