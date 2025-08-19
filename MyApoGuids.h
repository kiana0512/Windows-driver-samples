// MyApoGuids.h —— 只声明（不要 DEFINE_*）
#pragma once
#include <guiddef.h>    // GUID
#include <Propsys.h>    // PROPERTYKEY

// CLSID / 属性集 GUID —— 只声明
EXTERN_C const GUID CLSID_MyCompanyEfxApo;
EXTERN_C const GUID MYCOMPANY_APO_PROPSETID;

// 我们的 PROPERTYKEY —— 只声明
EXTERN_C const PROPERTYKEY PKEY_MyCompany_ParamsBlob;
EXTERN_C const PROPERTYKEY PKEY_MyCompany_Gain;
EXTERN_C const PROPERTYKEY PKEY_MyCompany_EQBand;
EXTERN_C const PROPERTYKEY PKEY_MyCompany_Reverb;
EXTERN_C const PROPERTYKEY PKEY_MyCompany_Limiter;

// 方案 C 的命名管道名
#define MYCOMPANY_PIPE_NAME  L"\\\\.\\pipe\\MyCompanyApoCtrl-USB_0A67_30A2_MI00"
