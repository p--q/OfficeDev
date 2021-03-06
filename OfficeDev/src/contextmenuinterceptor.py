#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui.ActionTriggerSeparatorType import LINE as ActionTriggerSeparatorType_LINE
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED, IGNORED
from com.sun.star.beans import UnknownPropertyException
from com.sun.star.lang import IndexOutOfBoundsException
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。
    if __name__ == "__main__":  # オートメーションのときはデバッグサーバーに接続しない。
        return func
    def wrapper(*args, **kwargs):
        import pydevd
        frame = None
        doc = XSCRIPTCONTEXT.getDocument()
        if doc:  # ドキュメントが取得できた時
            frame = doc.getCurrentController().getFrame()  # ドキュメントのフレームを取得。
        else:
            currentframe = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()  # モードレスダイアログのときはドキュメントが取得できないので、モードレスダイアログのフレームからCreatorのフレームを取得する。
            frame = currentframe.getCreator()
        if frame:   
            import time
            indicator = frame.createStatusIndicator()  # フレームからステータスバーを取得する。
            maxrange = 2  # ステータスバーに表示するプログレスバーの目盛りの最大値。2秒ロスするが他に適当な告知手段が思いつかない。
            indicator.start("Trying to connect to the PyDev Debug Server for about 20 seconds.", maxrange)  # ステータスバーに表示する文字列とプログレスバーの目盛りを設定。
            t = 1  # プレグレスバーの初期値。
            while t<=maxrange:  # プログレスバーの最大値以下の間。
                indicator.setValue(t)  # プレグレスバーの位置を設定。
                time.sleep(1)  # 1秒待つ。
                t += 1  # プログレスバーの目盛りを増やす。
            indicator.end()  # reset()の前にend()しておかないと元に戻らない。
            indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。
        pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。
        try:
            func(*args, **kwargs)  # Step Intoして中に入る。
        except:
            import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
    return wrapper
# @enableRemoteDebugging
def macro():  # マクロで実行するとフリーズする。
    doc = XSCRIPTCONTEXT.getDocument()  # マクロを起動した時のドキュメントのモデルを取得。 
    infomsg = "All context menus of the created document frame contains now a 'Help' entry with the submenus 'Content', 'Help Agent' and 'Tips'.\n\nPress 'Return' in the shell to remove the context menu interceptor and finish the example!"
    doc.getText().setString(infomsg)
    docframe = doc.getCurrentController().getFrame()
    doc.getCurrentController().getViewSettings().setPropertyValue("ZoomType", 0)
    controller = docframe.getController()
    contextmenuinterceptor = ContextMenuInterceptor()
    controller.registerContextMenuInterceptor(contextmenuinterceptor)
    print("\n ... all context menus of the created document frame contains now a 'Help' entry with the\n     submenus 'Content', 'Help Agent' and 'Tips'.\n\nPress 'Return' to remove the context menu interceptor and finish the example!")
    input()
    controller.releaseContextMenuInterceptor(contextmenuinterceptor)
    print(" ... context menu interceptor removed!")
    if hasattr(doc, "close"):
        doc.close(False)
    else:
        doc.dispose()
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):
    def notifyContextMenuExecute(self, contextmenuexecuteevent):
        try:
            contextmenu = contextmenuexecuteevent.ActionTriggerContainer
            rootmenuentry = contextmenu.createInstance("com.sun.star.ui.ActionTrigger")
            separator = contextmenu.createInstance("com.sun.star.ui.ActionTriggerSeparator") 
            separator.setPropertyValue("SeparatorType", ActionTriggerSeparatorType_LINE)
            submenucontainer = contextmenu.createInstance("com.sun.star.ui.ActionTriggerContainer") 
            rootmenuentry.setPropertyValue("Text", "Help")
            rootmenuentry.setPropertyValue("CommandURL", "slot:5410")
            rootmenuentry.setPropertyValue("HelpURL", "5410")
            rootmenuentry.setPropertyValue("SubContainer", submenucontainer)
            menuentry = contextmenu.createInstance("com.sun.star.ui.ActionTrigger") 
            menuentry.setPropertyValue("Text", "Content")
            menuentry.setPropertyValue("CommandURL", "slot:5401")
            menuentry.setPropertyValue("HelpURL", "5401")
            submenucontainer.insertByIndex(0, menuentry)
            menuentry = contextmenu.createInstance("com.sun.star.ui.ActionTrigger") 
            menuentry.setPropertyValue("Text", "Help Agent")
            menuentry.setPropertyValue("CommandURL", "slot:5962")
            menuentry.setPropertyValue("HelpURL", "5962")    
            submenucontainer.insertByIndex(1, menuentry)
            menuentry = contextmenu.createInstance("com.sun.star.ui.ActionTrigger") 
            menuentry.setPropertyValue("Text", "Tips")
            menuentry.setPropertyValue("CommandURL", "slot:5404")
            menuentry.setPropertyValue("HelpURL", "5404")      
            submenucontainer.insertByIndex(2, menuentry)
            contextmenu.insertByIndex(0, separator)
            contextmenu.insertByIndex(0, rootmenuentry)
            return EXECUTE_MODIFIED
        except UnknownPropertyException as e:
            print(e)
        except IndexOutOfBoundsException as e:    
            print(e)
        except Exception as e:    
            print(e)
        except:
            import traceback; traceback.print_exc()
        return IGNORED
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。
if __name__ == "__main__":  # オートメーションで実行するとき
    import officehelper
    from functools import wraps
    import sys
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
    sys.exit()  # これないとスクリプトがバッググラウンドで動いている。