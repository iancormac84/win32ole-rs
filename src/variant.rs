use std::ffi::c_void;

use windows::{
    core::{ManuallyDrop, PSTR},
    Win32::System::{
        Com::{
            SAFEARRAY, VARENUM, VARIANT, VT_ARRAY, VT_BYREF, VT_RECORD, VT_TYPEMASK, VT_VARIANT,
        },
        Ole::{
            IRecordInfo, SafeArrayGetDim, SafeArrayGetLBound, SafeArrayGetRecordInfo,
            SafeArrayGetUBound, SafeArrayLock, SafeArrayPtrOfIndex, SafeArrayUnlock,
        },
    },
};

fn ary_new_dim<'a, T>(myary: &'a mut Vec<T>, pid: &'a [i32], plb: &'a [i32], dim: u32) -> &'a mut Vec<T> {
    let ids: Vec<usize> = pid.iter().zip(plb).map(|(x, y)| (x - y) as usize).collect();

    let mut obj = myary;
    let mut pobj = myary;
    for i in 0..dim - 1 {
        obj = match pobj.get_mut(ids[i as usize]) {
            Some(inner_arr) => inner_arr,
            None => {
                let new_vec = Vec::new();
                pobj.insert(ids[i as usize], new_vec);
                pobj.get_mut(ids[i as usize]).unwrap()
            }
        };
        pobj = obj;
    }
    obj
}

fn ary_store_dim<T>(myary: &mut Vec<Vec<T>>, pid: &[i32], plb: &[i32], dim: u32, val: T) {
    let id = (pid[dim as usize - 1] - plb[dim as usize - 1]) as usize;
    let obj = ary_new_dim(myary, pid, plb, dim);
    obj.insert(id, val);
}

pub trait VariantAccess {
    fn vartype(&self) -> VARENUM;
    fn set_vartype(&mut self, vt: VARENUM);
    fn variant_ref(&self) -> *mut VARIANT;
    fn to_value(&mut self) -> T;
    fn is_array(&self) -> bool;
    fn is_byref(&self) -> bool;
    fn array_ref(&self) -> *mut *mut SAFEARRAY;
    fn array(&self) -> *mut SAFEARRAY;
    fn record_info(&self) -> ManuallyDrop<IRecordInfo>;
    fn set_record_info(&mut self, record_info: &IRecordInfo);
    fn record(&self) -> *mut c_void;
    fn byref(&self) -> *mut c_void;
    fn i1_ref(&self) -> PSTR;
}

impl VariantAccess for VARIANT {
    fn vartype(&self) -> VARENUM {
        unsafe { self.Anonymous.Anonymous.vt }
    }
    fn set_vartype(&mut self, vt: VARENUM) {
        unsafe {
            (*self.Anonymous.Anonymous).vt = vt;
        }
    }
    fn variant_ref(&self) -> *mut VARIANT {
        unsafe { self.Anonymous.Anonymous.Anonymous.pvarVal }
    }
    fn is_array(&self) -> bool {
        self.vartype().0 & VT_ARRAY.0 != 0
    }
    fn is_byref(&self) -> bool {
        self.vartype().0 & VT_BYREF.0 != 0
    }
    fn array_ref(&self) -> *mut *mut SAFEARRAY {
        unsafe { self.Anonymous.Anonymous.Anonymous.pparray }
    }
    fn array(&self) -> *mut SAFEARRAY {
        unsafe { self.Anonymous.Anonymous.Anonymous.parray }
    }
    fn record_info(&self) -> ManuallyDrop<IRecordInfo> {
        unsafe { self.Anonymous.Anonymous.Anonymous.Anonymous.pRecInfo }
    }
    fn set_record_info(&mut self, record_info: &IRecordInfo) {
        unsafe {
            (*(*self.Anonymous.Anonymous).Anonymous.Anonymous).pRecInfo = ManuallyDrop::new(record_info)
        };
    }
    fn record(&self) -> *mut c_void {
        unsafe { self.Anonymous.Anonymous.Anonymous.Anonymous.pvRecord }
    }
    fn byref(&self) -> *mut c_void {
        unsafe { self.Anonymous.Anonymous.Anonymous.byref }
    }
    fn i1_ref(&self) -> PSTR {
        unsafe { self.Anonymous.Anonymous.Anonymous.pcVal }
    }
    fn to_value(&mut self) -> Option<T> {
        let mut obj = None;
        let mut val = None;
        let mut vt = self.vartype();
        while vt.0 == VT_BYREF.0 | VT_VARIANT.0 {
            self = &mut unsafe{*self.variant_ref()};
            vt = self.vartype();
        }

        if self.is_array() {
            let vt_base = vt.0 & VT_TYPEMASK.0;
            let psa = if self.is_byref() {
                unsafe { *self.array_ref() }
            } else {
                self.array()
            };
            if psa.is_null() {
                return None;
            }
            let dim = unsafe { SafeArrayGetDim(psa) };
            let mut id = vec![0; dim as usize];
            let mut lb = vec![0; dim as usize];
            let mut ub = vec![0; dim as usize];
            for i in 0..dim {
                lb[i as usize] = unsafe { SafeArrayGetLBound(psa, i + 1).unwrap() };
                id[i as usize] = unsafe { SafeArrayGetLBound(psa, i + 1).unwrap() };
                ub[i as usize] = unsafe { SafeArrayGetUBound(psa, i + 1).unwrap() };
            }
            let result = unsafe { SafeArrayLock(psa) };
            if let Ok(()) = result {
                let mut obj = vec![];
                let mut i = 0;
                let mut variant = VARIANT::default();
                variant.set_vartype(VARENUM(vt_base | VT_BYREF.0));
                if vt_base == VT_RECORD.0 {
                    let record = unsafe { SafeArrayGetRecordInfo(psa) };
                    if let Ok(record) = record {
                        variant.set_vartype(VT_RECORD);
                        variant.set_record_info(&record);
                    }
                }
                while i < dim {
                    let obj = ary_new_dim(&mut obj, &id, &lb, dim);
                    let result = if vt_base == VT_RECORD.0 {
                        unsafe { SafeArrayPtrOfIndex(psa, id.as_ptr(), &mut variant.record()) }
                    } else {
                        unsafe { SafeArrayPtrOfIndex(psa, id.as_ptr(), &mut variant.byref()) }
                    };
                    if let Ok(()) = result {
                        val = variant.to_value();
                        ary_store_dim(obj, &id, &lb, dim, val);
                    }
                    for i in 0..dim as usize {
                        let new_pid = id[i] + 1;
                        id[i] = new_pid;
                        if id[i] <= ub[i] {
                            break;
                        }
                        id[i] = lb[i];
                    }
                }
                let result = unsafe { SafeArrayUnlock(psa) };
            }
            return obj;
        }
        let vt = self.vartype().0 & !VT_BYREF.0;
        match vt {
            VT_EMPTY => return None,
            VT_NULL => return None,
            VT_I1 => if self.is_byref() {},
        }
    }
}
