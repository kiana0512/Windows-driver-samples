// MyApoGuids.cpp —— 真正定义 GUID/PROPERTYKEY
#define INITGUID
#include <initguid.h>
#include <Propsys.h>
#include "MyApoGuids.h"

// 1) 你的 APO CLSID（可改成你自己的）
DEFINE_GUID(CLSID_MyCompanyEfxApo,
    0x8e3e0b71,0x5b8a,0x45c9,0x9b,0x3d,0x3a,0x2e,0x5b,0x41,0x8a,0x10);

// 2) 我们的“参数属性集” GUID 
DEFINE_GUID(MYCOMPANY_APO_PROPSETID,
    0xd4d9a040,0x8b5f,0x4c0e,0xaa,0xd1,0xaa,0xbb,0xcc,0xdd,0xee,0xff);

// 3) PROPERTYKEY（手动 {GUID,pid} 初始化）
const PROPERTYKEY PKEY_MyCompany_ParamsBlob = { MYCOMPANY_APO_PROPSETID, 10 };
const PROPERTYKEY PKEY_MyCompany_Gain       = { MYCOMPANY_APO_PROPSETID, 1  };
const PROPERTYKEY PKEY_MyCompany_EQBand     = { MYCOMPANY_APO_PROPSETID, 2  };
const PROPERTYKEY PKEY_MyCompany_Reverb     = { MYCOMPANY_APO_PROPSETID, 3  };
const PROPERTYKEY PKEY_MyCompany_Limiter    = { MYCOMPANY_APO_PROPSETID, 4  };
