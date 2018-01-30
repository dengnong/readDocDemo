import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * Created by 54472 on 2018/1/29.
 */
public class Jacob {

    /**
     *
     * JACOB一个Java-COM中间件.通过这个组件你可以在Java应用程序中调用COM组件和Win32 libraries。”
     * Jacob只能用于windows系统，如果你的系统不是windows，建议使用Openoffice.org
     * 1、到官网下载Jacob，目前最新版是1.17，地址链接：http://sourceforge.net/projects/jacob-project/
     * 2、将压缩包解压后，Jacob.jar添加到Libraries中(先复制到项目目录中，右键单击jar包选择Build Path—>Add to Build Path)；
     * 3、将Jacob.dll放至当前项目所用到的“jre\bin”下面(比如我的Eclipse正在用的Jre路径是D:\Java\jdk1.7.0_17\jre\bin)。
     * 可能还要将Jacob.dll放至“WINDOWS\SYSTEM32”下面
     */
    public static void DocToHtml(String filePath, String htmlPath) {
        ActiveXComponent app = new ActiveXComponent("Word.Application");
        try {
            app.setProperty("Visible", new Variant(false));
            Dispatch docs = app.getProperty("Documents").toDispatch();
            Dispatch doc = Dispatch.invoke(
                    docs,
                    "Open",
                    Dispatch.Method,
                    new Object[] { filePath, new Variant(false),
                            new Variant(true) }, new int[1])
                    .toDispatch();
            // 10代表word保存成筛选过的html
            Dispatch.invoke(
                    doc,
                    "SaveAs",
                    Dispatch.Method,
                    new Object[] {
                            htmlPath, new Variant(10) }, new int[1]);
            Dispatch.call(doc, "Close", new Variant(false));
        } catch (Exception e) {

        }
    }
}
