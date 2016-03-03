#pragma once

#include <atlbase.h>
#include<iostream>
#include <mshtml.h>
#include <winuser.h>
#include <comdef.h>
#include <string.h>
#include <atlcom.h>
#include<windows.h>
#include <atlstr.h>
#include "exdisp.h"

using namespace std;

void EnumIE(void);//处理网页
void EnumFrame(IHTMLDocument2 * pIHTMLDocument2);//处理框架
void EnumForm(IHTMLDocument2 * pIHTMLDocument2);//处理表单
CComModule _Module;  //使用CComDispatchDriver ATL的智能指针，此处必须声明

void EnumField(CComDispatchDriver spInputElement,CString ComType,CString ComVal,CString ComName);//处理表单域

void EnumIE(void)   
{   
   OleInitialize(NULL);//初始化com库  
   HRESULT   hr;  
   IWebBrowser2*    spBrowser;
   VARIANT vPostData;
   VariantInit(&vPostData);

    CoCreateInstance(CLSID_InternetExplorer, NULL, CLSCTX_LOCAL_SERVER, 
                       IID_IWebBrowser2, (void**)&spBrowser);
 
    if (spBrowser==NULL) return; 

	spBrowser->put_Visible(VARIANT_TRUE); //Commentout this line if you dont want the browser to be displayed

	VARIANT vEmpty;
    VariantInit(&vEmpty);
	VARIANT var;  
	var.vt = VT_BSTR;
	//var.bstrVal=CComBSTR(cIEUrl_Filter);
	var.bstrVal = SysAllocString(L"www.baidu.com");//改为校园网登录页网址

	hr=spBrowser->Navigate2(&var,&vEmpty,&vEmpty,&vEmpty,&vEmpty) ; //Open the URL page
	if (SUCCEEDED(hr))
    {
        spBrowser->put_Visible(VARIANT_TRUE);
    }
    else
    {
        spBrowser->Quit();
    }

	BOOL bReady=0;
	BSTR bsStatus;
	CString mStr;
	while(!bReady)  //This while loop maks sure that the page is fully loaded before we go to the next page
	{
		//如果用户手动关闭IE窗口,退出循环
		SHANDLE_PTR hHwnd;
		spBrowser->get_HWND(&hHwnd);
		if (NULL == hHwnd)
		{
			 bReady=1;
			 return;
		}

		//等待网页完全打开,退出循环
		spBrowser->get_StatusText(&bsStatus);
		mStr=bsStatus;
		if(mStr=="完毕" || mStr=="完成" || mStr=="Done" )
		{
			bReady=1;
		}

		Sleep(200);
	 }

	CComPtr<IDispatch> spDispDoc;   
	hr=spBrowser->get_Document(&spDispDoc);   
	if (FAILED(hr)) 
	{
		spBrowser->Release();
		OleUninitialize();
		return; 
	}
 
	CComQIPtr<IHTMLDocument2> spDocument2 =spDispDoc;   
	if (!spDocument2) 
	{
		spBrowser->Release();
		OleUninitialize();
		return;      
	}

	EnumForm(spDocument2); //枚举所有的表单
    
	spBrowser->Release();
	OleUninitialize();
}

void EnumFrame(IHTMLDocument2 * pIHTMLDocument2)
{   
 if (!pIHTMLDocument2) return;       
 HRESULT   hr;   
    
 CComPtr<IHTMLFramesCollection2> spFramesCollection2;   
 pIHTMLDocument2->get_frames(&spFramesCollection2); //取得框架frame的集合   
    
 long nFrameCount=0;        //取得子框架个数   
 hr=spFramesCollection2->get_length(&nFrameCount);   
 if (FAILED(hr)|| 0==nFrameCount) return;   
    
 for(long i=0; i<nFrameCount; i++)   
 {   
  CComVariant vDispWin2; //取得子框架的自动化接口   
  hr = spFramesCollection2->item(&CComVariant(i), &vDispWin2);   
  if (FAILED(hr)) continue;       
  CComQIPtr<IHTMLWindow2>spWin2 = vDispWin2.pdispVal;   
  if (!spWin2) continue; //取得子框架的   IHTMLWindow2   接口       
  CComPtr <IHTMLDocument2> spDoc2;   
  spWin2->get_document(&spDoc2); //取得子框架的   IHTMLDocument2   接口
  
  EnumForm(spDoc2);      //递归枚举当前子框架   IHTMLDocument2   上的表单form   
 }   
}

void EnumForm(IHTMLDocument2 * pIHTMLDocument2)   
{   
  if (!pIHTMLDocument2) return; 

  EnumFrame(pIHTMLDocument2);   //递归枚举当前IHTMLDocument2上的子框架frame  

  HRESULT hr;
      
  USES_CONVERSION;      

  CComQIPtr<IHTMLElementCollection> spElementCollection;   
  hr=pIHTMLDocument2->get_forms(&spElementCollection); //取得表单集合   
  if (FAILED(hr))   
  {   
    return;   
  }   
    
  long nFormCount=0;           //取得表单数目   
  hr=spElementCollection->get_length(&nFormCount);   
  if (FAILED(hr))   
  {   
    return;   
  }   
    
  for(long i=0; i<nFormCount; i++)   
  {   
    IDispatch *pDisp = NULL;   //取得第i项表单   
    hr=spElementCollection->item(CComVariant(i),CComVariant(),&pDisp);   
    if (FAILED(hr)) continue;   
    
    CComQIPtr<IHTMLFormElement> spFormElement= pDisp;   
    pDisp->Release();   
    
    long nElemCount=0;         //取得表单中域的数目   
    hr=spFormElement->get_length(&nElemCount);   
    if (FAILED(hr)) continue;   
    
    for(long j=0; j<nElemCount; j++)   
	{   
    
      CComDispatchDriver spInputElement; //取得第j项表单域   
      hr=spFormElement->item(CComVariant(j), CComVariant(), &spInputElement);   
      if (FAILED(hr)) continue;  

      CComVariant vName,vVal,vType;     //取得表单域的名称，数值，类型 
      hr=spInputElement.GetPropertyByName(L"name", &vName);   
      if (FAILED(hr)) continue;   
      hr=spInputElement.GetPropertyByName(L"value", &vVal);   
      if(FAILED(hr)) continue;   
      hr=spInputElement.GetPropertyByName(L"type", &vType);   
      if(FAILED(hr)) continue;   
    
      LPCTSTR lpName= vName.bstrVal ? OLE2CT(vName.bstrVal) : _T("NULL"); //未知域名   
      LPCTSTR lpVal=  vVal.bstrVal  ? OLE2CT(vVal.bstrVal)  : _T("NULL"); //空值，未输入   
      LPCTSTR lpType= vType.bstrVal ? OLE2CT(vType.bstrVal) : _T("NULL"); //未知类型  
   
	  EnumField(spInputElement,lpType,lpVal,lpName);//传递并处理表单域的类型、值、名
	}//表单域循环结束     
  }//表单循环结束   
}  

void EnumField(CComDispatchDriver spInputElement,CString ComType,CString ComVal,CString ComName)
{//处理表单域
	if ((ComType.Find("text")>=0) && ComVal.Compare(CString("NULL"))==0 && ComName.Compare(CString("username"))==0)
   { 
        CString Tmp="";//add your username
        CComVariant vSetStatus(Tmp);
        spInputElement.PutPropertyByName(L"value",&vSetStatus);
   }
   if ((ComType.Find("password")>=0) && ComVal.Compare(CString("NULL"))==0&& ComName.Compare(CString("password"))==0)
   { 
        CString Tmp="";//add your password
        CComVariant vSetStatus(Tmp);
	    spInputElement.PutPropertyByName(L"value",&vSetStatus);
   }
   if ((ComType.Find("submit")>=0))
   {
		IHTMLElement*  pHElement;
		spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pHElement);
		pHElement->click();                
   }
}
