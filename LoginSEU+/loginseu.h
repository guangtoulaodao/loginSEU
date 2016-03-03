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

void EnumIE(void);//������ҳ
void EnumFrame(IHTMLDocument2 * pIHTMLDocument2);//������
void EnumForm(IHTMLDocument2 * pIHTMLDocument2);//�����
CComModule _Module;  //ʹ��CComDispatchDriver ATL������ָ�룬�˴���������

void EnumField(CComDispatchDriver spInputElement,CString ComType,CString ComVal,CString ComName);//�������

void EnumIE(void)   
{   
   OleInitialize(NULL);//��ʼ��com��  
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
	var.bstrVal = SysAllocString(L"www.baidu.com");//��ΪУ԰����¼ҳ��ַ

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
		//����û��ֶ��ر�IE����,�˳�ѭ��
		SHANDLE_PTR hHwnd;
		spBrowser->get_HWND(&hHwnd);
		if (NULL == hHwnd)
		{
			 bReady=1;
			 return;
		}

		//�ȴ���ҳ��ȫ��,�˳�ѭ��
		spBrowser->get_StatusText(&bsStatus);
		mStr=bsStatus;
		if(mStr=="���" || mStr=="���" || mStr=="Done" )
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

	EnumForm(spDocument2); //ö�����еı�
    
	spBrowser->Release();
	OleUninitialize();
}

void EnumFrame(IHTMLDocument2 * pIHTMLDocument2)
{   
 if (!pIHTMLDocument2) return;       
 HRESULT   hr;   
    
 CComPtr<IHTMLFramesCollection2> spFramesCollection2;   
 pIHTMLDocument2->get_frames(&spFramesCollection2); //ȡ�ÿ��frame�ļ���   
    
 long nFrameCount=0;        //ȡ���ӿ�ܸ���   
 hr=spFramesCollection2->get_length(&nFrameCount);   
 if (FAILED(hr)|| 0==nFrameCount) return;   
    
 for(long i=0; i<nFrameCount; i++)   
 {   
  CComVariant vDispWin2; //ȡ���ӿ�ܵ��Զ����ӿ�   
  hr = spFramesCollection2->item(&CComVariant(i), &vDispWin2);   
  if (FAILED(hr)) continue;       
  CComQIPtr<IHTMLWindow2>spWin2 = vDispWin2.pdispVal;   
  if (!spWin2) continue; //ȡ���ӿ�ܵ�   IHTMLWindow2   �ӿ�       
  CComPtr <IHTMLDocument2> spDoc2;   
  spWin2->get_document(&spDoc2); //ȡ���ӿ�ܵ�   IHTMLDocument2   �ӿ�
  
  EnumForm(spDoc2);      //�ݹ�ö�ٵ�ǰ�ӿ��   IHTMLDocument2   �ϵı�form   
 }   
}

void EnumForm(IHTMLDocument2 * pIHTMLDocument2)   
{   
  if (!pIHTMLDocument2) return; 

  EnumFrame(pIHTMLDocument2);   //�ݹ�ö�ٵ�ǰIHTMLDocument2�ϵ��ӿ��frame  

  HRESULT hr;
      
  USES_CONVERSION;      

  CComQIPtr<IHTMLElementCollection> spElementCollection;   
  hr=pIHTMLDocument2->get_forms(&spElementCollection); //ȡ�ñ�����   
  if (FAILED(hr))   
  {   
    return;   
  }   
    
  long nFormCount=0;           //ȡ�ñ���Ŀ   
  hr=spElementCollection->get_length(&nFormCount);   
  if (FAILED(hr))   
  {   
    return;   
  }   
    
  for(long i=0; i<nFormCount; i++)   
  {   
    IDispatch *pDisp = NULL;   //ȡ�õ�i���   
    hr=spElementCollection->item(CComVariant(i),CComVariant(),&pDisp);   
    if (FAILED(hr)) continue;   
    
    CComQIPtr<IHTMLFormElement> spFormElement= pDisp;   
    pDisp->Release();   
    
    long nElemCount=0;         //ȡ�ñ��������Ŀ   
    hr=spFormElement->get_length(&nElemCount);   
    if (FAILED(hr)) continue;   
    
    for(long j=0; j<nElemCount; j++)   
	{   
    
      CComDispatchDriver spInputElement; //ȡ�õ�j�����   
      hr=spFormElement->item(CComVariant(j), CComVariant(), &spInputElement);   
      if (FAILED(hr)) continue;  

      CComVariant vName,vVal,vType;     //ȡ�ñ�������ƣ���ֵ������ 
      hr=spInputElement.GetPropertyByName(L"name", &vName);   
      if (FAILED(hr)) continue;   
      hr=spInputElement.GetPropertyByName(L"value", &vVal);   
      if(FAILED(hr)) continue;   
      hr=spInputElement.GetPropertyByName(L"type", &vType);   
      if(FAILED(hr)) continue;   
    
      LPCTSTR lpName= vName.bstrVal ? OLE2CT(vName.bstrVal) : _T("NULL"); //δ֪����   
      LPCTSTR lpVal=  vVal.bstrVal  ? OLE2CT(vVal.bstrVal)  : _T("NULL"); //��ֵ��δ����   
      LPCTSTR lpType= vType.bstrVal ? OLE2CT(vType.bstrVal) : _T("NULL"); //δ֪����  
   
	  EnumField(spInputElement,lpType,lpVal,lpName);//���ݲ������������͡�ֵ����
	}//����ѭ������     
  }//��ѭ������   
}  

void EnumField(CComDispatchDriver spInputElement,CString ComType,CString ComVal,CString ComName)
{//�������
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
