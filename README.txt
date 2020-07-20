[20200719-08:30]
OleDataSource项目含有
允许剪切板传输数据过程中实现回调的COM类以及其它辅助类。

如何测试？
运行后出现一个空白窗口，点击“帮助”->“关于”菜单，不会弹出关于对话框，
而是将包含test.txt文件的路径的COM对象写入剪切板。
然后在桌面上右键粘贴即可。

哪些类和文件最重要？
为了结构完整，多加了很多类，但是多数都用不到。
最重要的有，
Usage.cpp : 
	HGLOBAL PutIntoClipboard(const TCHAR* path)

CMyOleDataSource.cpp : 
	BOOL CMyOleDataSource::OnRenderData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium)

CMyOleStream.cpp:
	HRESULT __stdcall CMyOleStream::Read(void* pv, ULONG cb, ULONG* pcbRead)
	HRESULT __stdcall CMyOleStream::Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition)

设计上有何改动？

原来的方法，是在OnRenderDataFile中创建了一个内存共享文件。对于小文件可以都
留在内存中然后一起写出，但是对大文件不能这样做。
现在的方法是，在OnRenderData中返回一个IStream对象给调用方（比如explorer），
调用方随后会调用IStream::Read（配合Seek获得文件长度），也就是在OnRenderData
这一次回调之后，对于IStream::Read进行二次回调。

这时候在CMyOleStream中自己实现这个回调，提供需要的数据即可。
这就可以实现边传输边提供数据给调用方的效果。

需要注意的问题:
在主消息循环之前要先初始化OLE

    // Initialize COM and OLE
    if (OleInitialize(0) != S_OK)
        return 0;
		
在之后要卸载OLE
    // Cleanup
    OleUninitialize();

详见WinLoop.cpp
