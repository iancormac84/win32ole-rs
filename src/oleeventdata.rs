use std::ptr;

use windows::{
    core::{implement, Interface, Vtable, BSTR, GUID, HSTRING},
    Win32::{
        Foundation::{DISP_E_BADINDEX, E_NOINTERFACE, HWND},
        Globalization::GetUserDefaultLCID,
        System::{
            Com::{
                IConnectionPoint, IConnectionPointContainer, IDispatch, IDispatch_Impl, ITypeInfo,
                IMPLTYPEFLAGS, IMPLTYPEFLAG_FDEFAULT, IMPLTYPEFLAG_FSOURCE, TKIND_COCLASS,
                TYPEATTR,
            },
            Ole::{IProvideClassInfo, IProvideClassInfo2, GUIDKIND_DEFAULT_SOURCE_DISP_IID},
        },
        UI::WindowsAndMessaging::{
            DispatchMessageW, PeekMessageW, TranslateMessage, MSG, PM_REMOVE,
        },
    },
};

use crate::{error::Result, OleData};

pub struct IEventSinkObject {
    event_sink: IEventSink,
    m_ref: u32,
    m_iid: GUID,
    m_event_id: u64,
    typeinfo: ITypeInfo,
}

pub struct OleEventData {
    cookie: u32,
    connection_point: IConnectionPoint,
    dispatch: IDispatch,
    event_id: u64,
}

impl Drop for OleEventData {
    fn drop(&mut self) {
        unsafe { self.connection_point.Unadvise(self.cookie) };
    }
}

fn ole_msg_loop() {
    let mut msg = MSG::default();
    unsafe {
        while PeekMessageW(&mut msg, HWND(0), 0, 0, PM_REMOVE).as_bool() {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
}

#[derive(Debug)]
struct GuidInfo {
    pub guid: Option<GUID>,
    pub typeinfo: Option<ITypeInfo>,
}

fn find_iid(oledata: &OleData, pitf: Option<&str>, piid: &GUID) -> Result<GuidInfo> {
    let typeinfo = unsafe { oledata.dispatch.GetTypeInfo(0, GetUserDefaultLCID()) }?;

    let mut typelib = None;
    let mut index = 0;
    unsafe { typeinfo.GetContainingTypeLib(&mut typelib, &mut index) }?;
    let typelib = typelib.unwrap();

    if pitf.is_none() {
        return Ok(GuidInfo {
            guid: None,
            typeinfo: Some(unsafe { typelib.GetTypeInfoOfGuid(piid)? }),
        });
    }
    let pitf = pitf.unwrap();
    let count = unsafe { typelib.GetTypeInfoCount() };
    for index in 0..count {
        let typeinfo = unsafe { typelib.GetTypeInfo(index) };
        let Ok(typeinfo) = typeinfo else {
            break;
        };
        let type_attr = unsafe { typeinfo.GetTypeAttr() };

        let Ok(type_attr) = type_attr else {
            break;
        };

        if unsafe { (*type_attr).typekind } == TKIND_COCLASS {
            for type_ in 0..unsafe { (*type_attr).cImplTypes } {
                let ref_type = unsafe { typeinfo.GetRefTypeOfImplType(type_ as u32) };
                let Ok(ref_type) = ref_type else {
                    break;
                };
                let impl_type_info = unsafe { typeinfo.GetRefTypeInfo(ref_type) };
                let Ok(impl_type_info) = impl_type_info else {
                    break;
                };

                let mut bstr = BSTR::default();
                let result = unsafe {
                    impl_type_info.GetDocumentation(
                        -1,
                        Some(&mut bstr),
                        None,
                        ptr::null_mut(),
                        None,
                    )
                };
                if result.is_err() {
                    break;
                }

                if pitf == bstr {
                    let impl_type_attr = unsafe { impl_type_info.GetTypeAttr() };
                    if let Ok(impl_type_attr) = impl_type_attr {
                        unsafe { typeinfo.ReleaseTypeAttr(type_attr) };
                        unsafe { impl_type_info.ReleaseTypeAttr(impl_type_attr) };
                        return Ok(GuidInfo {
                            guid: Some(unsafe { (*impl_type_attr).guid }),
                            typeinfo: Some(impl_type_info),
                        });
                    } else {
                        break;
                    }
                }
            }
        }
        unsafe { typeinfo.ReleaseTypeAttr(type_attr) };
    }
    let msg = HSTRING::from(format!("failed to find GUID or ITypeInfo for {pitf}"));
    Err(windows::core::Error::new(E_NOINTERFACE, msg).into())
}

struct ITypeInfoData<'a> {
    pub typeinfo: ITypeInfo,
    pub typedata: &'a TYPEATTR,
}

impl Drop for ITypeInfoData<'_> {
    fn drop(&mut self) {
        unsafe { self.typeinfo.ReleaseTypeAttr(self.typedata) };
    }
}

#[implement(IDispatch)]
pub struct IEventSink();

impl IDispatch_Impl for IEventSink {
    fn GetTypeInfoCount(&self) -> windows::core::Result<u32> {
        Ok(0)
    }

    fn GetTypeInfo(&self, _itinfo: u32, _lcid: u32) -> windows::core::Result<ITypeInfo> {
        Err(DISP_E_BADINDEX.into())
    }

    fn GetIDsOfNames(
        &self,
        riid: *const windows::core::GUID,
        rgsznames: *const windows::core::PCWSTR,
        cnames: u32,
        lcid: u32,
        rgdispid: *mut i32,
    ) -> windows::core::Result<()> {
        todo!()
    }

    fn Invoke(
        &self,
        dispidmember: i32,
        riid: *const windows::core::GUID,
        lcid: u32,
        wflags: windows::Win32::System::Com::DISPATCH_FLAGS,
        pdispparams: *const windows::Win32::System::Com::DISPPARAMS,
        pvarresult: *mut windows::Win32::System::Com::VARIANT,
        pexcepinfo: *mut windows::Win32::System::Com::EXCEPINFO,
        puargerr: *mut u32,
    ) -> windows::core::Result<()> {
        todo!()
    }
}

/*fn ev_advise(oledata: &OleData, itf: Option<&str>) -> Result<()> {
    let guid_info = if itf.is_some() {
        //Creation of piid is just to appease the function signature's demands
        let piid = GUID::new().unwrap();
        find_iid(oledata, itf, &piid)
    } else {
        find_default_source(oledata)
    };
    let Ok(guid_info) = guid_info else {
        return Err(guid_info.unwrap_err().into());
    };
    let connection_point_container = ptr::null();
    let result = unsafe { oledata.dispatch.query(&IConnectionPointContainer::IID, &mut connection_point_container) };
    let result = result.ok();
    if let Err(error) = result {
        return Err(error.into());
    }
    let connection_point_container = unsafe { <IConnectionPointContainer as Vtable>::from_raw(connection_point_container as *mut _) };
    let connection_point = unsafe { connection_point_container.FindConnectionPoint(&guid_info.guid.unwrap()) }?;
}*/

fn find_coclass<'a>(typeinfo: &ITypeInfo, typeattr: &TYPEATTR) -> Result<ITypeInfoData<'a>> {
    let mut typelib = None;
    unsafe { typeinfo.GetContainingTypeLib(&mut typelib, ptr::null_mut()) }?;
    let typelib = typelib.unwrap();
    let count = unsafe { typelib.GetTypeInfoCount() };
    for i in 0..count {
        let typeinfo2 = unsafe { typelib.GetTypeInfo(i) };
        let Ok(typeinfo2) = typeinfo2 else {
            continue;
        };
        let typeattr2 = unsafe { typeinfo2.GetTypeAttr() };
        let Ok(typeattr2) = typeattr2 else {
            continue;
        };
        if unsafe { (*typeattr2).typekind } != TKIND_COCLASS {
            unsafe { typeinfo2.ReleaseTypeAttr(typeattr2) };
            continue;
        }
        for j in 0..unsafe { (*typeattr2).cImplTypes } {
            let flags = unsafe { typeinfo2.GetImplTypeFlags(j as u32) };
            let Ok(flags) = flags else {
                continue;
            };
            if flags & IMPLTYPEFLAG_FDEFAULT == IMPLTYPEFLAGS(0) {
                continue;
            }
            let href = unsafe { typeinfo2.GetRefTypeOfImplType(j as u32) };
            let Ok(href) = href else {
                continue;
            };
            let reftypeinfo = unsafe { typeinfo2.GetRefTypeInfo(href) };
            let Ok(reftypeinfo) = reftypeinfo else {
                continue;
            };
            let reftypeattr = unsafe { reftypeinfo.GetTypeAttr() };
            let Ok(reftypeattr) = reftypeattr else {
                continue;
            };
            if typeattr.guid == unsafe { (*reftypeattr).guid } {
                return Ok(ITypeInfoData {
                    typeinfo: typeinfo2,
                    typedata: unsafe { &*typeattr2 },
                });
            } else {
                unsafe { typeinfo2.ReleaseTypeAttr(typeattr2) };
            }
        }
    }
    let msg = HSTRING::from(format!(
        "failed to find ITypeInfoData for {:?}",
        typeattr.guid
    ));
    Err(windows::core::Error::new(E_NOINTERFACE, msg).into())
}

fn find_default_source_from_typeinfo(
    typeinfo: &ITypeInfo,
    type_attr: &TYPEATTR,
) -> Result<ITypeInfo> {
    /* Enumerate all implemented types of the COCLASS */
    let mut result = Ok(());
    let mut ret_type_info = None;
    for i in 0..type_attr.cImplTypes {
        let flags = unsafe { typeinfo.GetImplTypeFlags(i as u32) };
        let Ok(flags) = flags else {
            continue;
        };

        /*
           looking for the [default] [source]
           we just hope that it is a dispinterface :-)
        */
        if (flags & IMPLTYPEFLAG_FDEFAULT) != IMPLTYPEFLAGS(0)
            && (flags & IMPLTYPEFLAG_FSOURCE) != IMPLTYPEFLAGS(0)
        {
            let hreftype = unsafe { typeinfo.GetRefTypeOfImplType(i as u32) };
            let Ok(hreftype) = hreftype else {
                continue;
            };
            let reftypeinfo = unsafe { typeinfo.GetRefTypeInfo(hreftype) };
            if let Err(error) = reftypeinfo {
                result = Err(error);
            } else {
                let ref_type_info = reftypeinfo.unwrap();
                ret_type_info = Some(ref_type_info);
                break;
            }
        }
    }
    if let Err(error) = result {
        Err(error.into())
    } else {
        Ok(ret_type_info.unwrap())
    }
}

fn find_default_source(ole: &OleData) -> Result<GuidInfo> {
    let mut provider_class_info_interface = ptr::null();
    let result = unsafe {
        ole.dispatch
            .query(&IProvideClassInfo2::IID, &mut provider_class_info_interface)
    };
    if result.is_ok() {
        let provide_class_info2 = unsafe {
            <IProvideClassInfo2 as Vtable>::from_raw(provider_class_info_interface as *mut _)
        };
        let piid =
            unsafe { provide_class_info2.GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID.0 as u32) };
        if let Ok(piid) = piid {
            let result = find_iid(ole, None, &piid);
            if let Ok(guidinfo) = result {
                return Ok(guidinfo);
            }
        }
    }

    let mut typeinfo = None;
    let result = unsafe {
        ole.dispatch
            .query(&IProvideClassInfo::IID, &mut provider_class_info_interface)
    };
    if result.is_ok() {
        let provide_class_info = unsafe {
            <IProvideClassInfo as Vtable>::from_raw(provider_class_info_interface as *mut _)
        };
        let classinfo = unsafe { provide_class_info.GetClassInfo() };
        if let Ok(classinfo) = classinfo {
            typeinfo = Some(classinfo);
        }
    }
    if typeinfo.is_none() {
        typeinfo = Some(unsafe { ole.dispatch.GetTypeInfo(0, GetUserDefaultLCID()) }?);
    }
    let typeinfo = typeinfo.unwrap();
    let typeattr = unsafe { typeinfo.GetTypeAttr() }?;
    let copair = ITypeInfoData {
        typeinfo,
        typedata: unsafe { &*typeattr },
    };

    let mut pptypeinfo = find_default_source_from_typeinfo(&copair.typeinfo, copair.typedata);
    if pptypeinfo.is_err() {
        let coclass = find_coclass(&copair.typeinfo, copair.typedata);
        if let Ok(coclass) = coclass {
            pptypeinfo = find_default_source_from_typeinfo(&coclass.typeinfo, coclass.typedata);
        }
    }

    let pptypeinfo = pptypeinfo?;

    /* Determine IID of default source interface */
    let pptypeattr = unsafe { pptypeinfo.GetTypeAttr() }?;
    let piid = unsafe { (*pptypeattr).guid };
    unsafe { pptypeinfo.ReleaseTypeAttr(pptypeattr) };
    Ok(GuidInfo {
        guid: Some(piid),
        typeinfo: Some(pptypeinfo),
    })
}
