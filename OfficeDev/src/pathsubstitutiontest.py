#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
import sys
from com.sun.star.container import NoSuchElementException
def main(ctx, smgr):  # ctx: コンポーネントコンテクスト、smgr: サービスマネジャー
	pathvals = "$(home)",\
				"$(inst)",\
				"$(instpath)",\
				"$(insturl)",\
				"$(prog)",\
				"$(progpath)",\
				"$(progurl)",\
				"$(temp)",\
				"$(user)",\
				"$(userpath)",\
				"$(userurl)",\
				"$(username)",\
				"$(work)",\
				"$(path)",\
				"$(lang)",\
				"$(langid)",\
				"$(vlang)"
	pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
	for pathval in pathvals:
		try:
			url = pathsubstservice.getSubstituteVariableValue(pathval)
		except NoSuchElementException:
			print("{} is unknown variable!".format(pathval), file=sys.stderr)
			continue
		systempath = os.path.normpath(unohelper.fileUrlToSystemPath(url))  # urlをシステムパスにして正規化する。
		print("{} : {}".format(pathval, systempath))	
	print("\nCheck the resubstitution function")
	path = pathsubstservice.getSubstituteVariableValue(pathvals[0])
	path += "/test"
	print("Path = {}".format(path))
	resustpath = pathsubstservice.reSubstituteVariables(path)
	print("Resubstituted path = {}".format(resustpath))
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