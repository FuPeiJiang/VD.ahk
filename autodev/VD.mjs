/**
 * VD for AutoDev (QuickJS JavaScript, x64 Windows 11).
 * Invoke with `autodev run --module <entry.mjs>`.
 * Windows desktop COM is vtable-only; VDNative.dll is the narrow C ABI adapter.
 * There is no AutoHotkey or PowerShell dependency.
 */

function libraryDirectory() {
  // Runtime-resolved script path works in file execution and packaged projects.
  const source = autodev.script?.path;
  if (typeof source !== 'string' || !source) {
    throw new Error('Specify dllPath when evaluating interactively (autodev.script.path is unset)');
  }
  const last = Math.max(source.lastIndexOf('\\'), source.lastIndexOf('/'));
  return source.slice(0, last + 1);
}
function normalizeNumber(value, name) {
  if (!Number.isInteger(value)) throw new RangeError(`${name} must be an integer: ${value}`);
  return value;
}
function checked(hr, operation) {
  if (hr !== 0) {
    const code = (hr >>> 0).toString(16).toUpperCase().padStart(8, '0');
    throw new Error(`VD.${operation} failed: HRESULT 0x${code}`);
  }
}

/** Stateful object owns two Win32 DLL handles and one VDNative.dll handle. */
export class VirtualDesktops {
  constructor({ dllPath } = {}) {
    this.dllPath = dllPath ?? `${libraryDirectory()}VDNative.dll`;
    this.dll = ffi.open(this.dllPath);
    this.user32 = ffi.open('user32.dll');
    this.kernel32 = ffi.open('kernel32.dll');
    this.functions = new Map();
    this.sleep = this.kernel32.bind('Sleep', {args:['u32'], returns:'void'});
    this.foreground = this.user32.bind('GetForegroundWindow', {returns:'pointer'});
    this.focus = this.user32.bind('SetForegroundWindow', {args:['pointer'], returns:'bool'});
    this.closed = false;
  }
  _fn(name, args) {
    if (this.closed) throw new Error('VD is already closed');
    if (!this.functions.has(name)) {
      this.functions.set(name, this.dll.bind(name, {args, returns:'i32'}));
    }
    return this.functions.get(name);
  }
  _out(name, args = [], values = []) {
    const out = ffi.alloc(4);
    try {
      checked(this._fn(name, [...args, 'pointer'])(...values, out), name);
      return out.readI32(0);
    } finally { out.free(); }
  }
  _call(name, args, values) {
    checked(this._fn(name, args)(...values), name);
  }
  _validateIndex(index) {
    normalizeNumber(index, 'desktop number');
    const count = this.getCount();
    if (index < 1 || index > count) throw new RangeError(`Desktop ${index} outside 1..${count}`);
    return index;
  }
  /** Queries the OS afresh; no cached raw COM interface survives Explorer restart. */
  getCount() { return this._out('vd_count'); }
  getCurrentDesktopNum() { return this._out('vd_current'); }
  getVersion() {
    const b = ffi.alloc(4), r = ffi.alloc(4), a = ffi.alloc(4);
    try {
      this._call('vd_version', ['pointer','pointer','pointer'], [b,r,a]);
      return {build:b.readI32(0),revision:r.readI32(0),abi:a.readI32(0)};
    } finally { a.free();r.free();b.free(); }
  }
  modulusResolveDesktopNum(n) {
    normalizeNumber(n,'desktop number');
    const count=this.getCount();
    if (!count) throw new Error('No virtual desktops available');
    return ((n - 1) % count + count) % count + 1;
  }
  goToDesktopNum(target, {wait = true, timeoutMs = 4000} = {}) {
    this._validateIndex(target);
    if (target !== this.getCurrentDesktopNum()) this._call('vd_switch', ['i32'], [target]);
    if (wait && !this.waitDesktopSwitched(target, timeoutMs)) {
      throw new Error(`Timed out switching to desktop ${target}`);
    }
    return target;
  }
  goToRelativeDesktopNum(offset, options) {
    normalizeNumber(offset,'relative desktop offset');
    return this.goToDesktopNum(this.modulusResolveDesktopNum(this.getCurrentDesktopNum()+offset),options);
  }
  waitDesktopSwitched(target, timeoutMs = 4000) {
    this._validateIndex(target);
    if (!Number.isFinite(timeoutMs) || timeoutMs < 0 || timeoutMs > 120000) throw new RangeError('timeoutMs out of range');
    const start = Date.now();
    do {
      if (this.getCurrentDesktopNum() === target) return true;
      if (Date.now()-start >= timeoutMs) break;
      this.sleep(25);
    } while (true);
    return false;
  }
  createDesktop({switchTo = false} = {}) {
    const number = this._out('vd_create');
    if (switchTo) this.goToDesktopNum(number);
    return number;
  }
  createUntil(count,{switchToLast=false}={}) {
    normalizeNumber(count,'desktop count');
    if (count<1 || count>100) throw new RangeError('Requested desktop count must be 1..100');
    let current=this.getCount();
    while(current<count) {
      const next=this.createDesktop();
      if(next<=current) throw new Error('Creation did not advance desktop count');
      current=next;
    }
    if(switchToLast && current===count) this.goToDesktopNum(current);
    return current;
  }
  removeDesktop(target = this.getCurrentDesktopNum(), fallback = 0) {
    this._validateIndex(target);
    if (this.getCount() === 1) throw new Error('Cannot remove the last virtual desktop');
    if (fallback!==0) this._validateIndex(fallback);
    if (target===fallback) throw new RangeError('Fallback cannot be the removed desktop');
    this._call('vd_remove', ['i32','i32'], [target,fallback]);
    return this.getCount();
  }
  /** HWND number, bigint, {hwnd}, 'A' (foreground), or exact/partial title. */
  resolveWindow(target) {
    if (typeof target==='bigint') return target;
    if (typeof target==='number' && Number.isSafeInteger(target) && target>0) return BigInt(target);
    if (target && typeof target==='object' && target.hwnd) return this.resolveWindow(target.hwnd);
    if (target==='A') {
      const hwnd=this.foreground();
      if (!hwnd || hwnd===0n) throw new Error('No foreground window');
      return hwnd;
    }
    if (typeof target==='string' && target) {
      const windows=autodev.windows.list().filter(w=>
        w.hwnd && (w.title===target || w.title?.includes(target)));
      if (windows.length===0) throw new Error(`No window matches ${JSON.stringify(target)}`);
      if (windows.length>1) throw new Error(`Ambiguous window ${JSON.stringify(target)} (${windows.length} matches); use HWND`);
      return BigInt(windows[0].hwnd);
    }
    throw new TypeError('Window target must be HWND, {hwnd}, title, or A');
  }
  getDesktopNumOfWindow(target) {
    return this._out('vd_window_desktop', ['pointer'], [this.resolveWindow(target)]);
  }
  moveWindowToDesktopNum(target,number,{follow=false}={}) {
    this._validateIndex(number);
    const hwnd=this.resolveWindow(target);
    this._call('vd_move_window',['pointer','i32'],[hwnd,number]);
    if (follow) {
      this.goToDesktopNum(number);
      this.focus(hwnd); // Windows may deny foreground permission; do not falsify success.
    }
    return number;
  }
  moveWindowToRelativeDesktopNum(target,offset,options) {
    normalizeNumber(offset,'relative desktop offset');
    return this.moveWindowToDesktopNum(target,this.modulusResolveDesktopNum(this.getCurrentDesktopNum()+offset),options);
  }
  moveWindowToCurrentDesktop(target,options) {
    return this.moveWindowToDesktopNum(target,this.getCurrentDesktopNum(),options);
  }
  goToDesktopOfWindow(target,{activate=true}={}) {
    const hwnd=this.resolveWindow(target);
    const number=this.getDesktopNumOfWindow(hwnd);
    // -1 = pinned app; -2 = pinned window. Both already live on all desktops.
    if(number>0) this.goToDesktopNum(number);
    if(activate) this.focus(hwnd);
    return number;
  }
  isWindowPinned(target) {
    return this._out('vd_is_pinned',['pointer'],[this.resolveWindow(target)])===1;
  }
  pinWindow(target) { this._call('vd_pin_window',['pointer','i32'],[this.resolveWindow(target),1]); }
  unpinWindow(target) { this._call('vd_pin_window',['pointer','i32'],[this.resolveWindow(target),0]); }
  togglePinWindow(target) { this.isWindowPinned(target)?this.unpinWindow(target):this.pinWindow(target); }
  isExePinned(appId) { return this._out('vd_is_app_pinned',['wstring'],[appId])===1; }
  pinExe(appId) { this._call('vd_pin_app',['wstring','i32'],[appId,1]); }
  unpinExe(appId) { this._call('vd_pin_app',['wstring','i32'],[appId,0]); }
  togglePinExe(appId) { this.isExePinned(appId)?this.unpinExe(appId):this.pinExe(appId); }
  /** kind 0 = desktop name, 1 = per-desktop wallpaper; UTF-16 output. */
  _getText(index,kind) {
    this._validateIndex(index);
    const buf=ffi.alloc(8192);
    try {
      this._call('vd_get_text',['i32','i32','pointer','i32'],[index,kind,buf,4096]);
      const bytes=buf.readBytes(0,8192);let text='';
      for(let i=0;i+1<bytes.length;i+=2) {
        const code=bytes[i]|(bytes[i+1]<<8);
        if(code===0)break;
        text+=String.fromCharCode(code);
      }
      return text;
    } finally {buf.free();}
  }
  getNameFromDesktopNum(index) { return this._getText(index,0); }
  getWallpaperFromDesktopNum(index) { return this._getText(index,1); }
  setNameToDesktopNum(index,name) {
    this._validateIndex(index);
    if (typeof name!=='string'||name.length>2048) throw new TypeError('Invalid desktop name');
    this._call('vd_set_text',['i32','i32','wstring'],[index,0,name]);
  }
  setWallpaperToDesktopNum(index,path) {
    this._validateIndex(index);
    if (typeof path!=='string'||path.length>2048) throw new TypeError('Invalid wallpaper path');
    this._call('vd_set_text',['i32','i32','wstring'],[index,1,path]);
  }
  /** Read-only snapshot, useful for inspection and regressions. */
  getInfo() {
    const current=this.getCurrentDesktopNum(), count=this.getCount();
    const desktops=[];
    for(let n=1;n<=count;n++) {
      let name=null;
      try { name=this.getNameFromDesktopNum(n); } catch(_ignored) { /* older Windows ABI */ }
      desktops.push({number:n,name,isCurrent:n===current});
    }
    return {version:this.getVersion(),count,current,desktops};
  }
  close() {
    if (this.closed) return;
    this.closed=true;
    this.functions.clear();
    this.user32.close();this.kernel32.close();this.dll.close();
  }
}

export const VD=VirtualDesktops;
