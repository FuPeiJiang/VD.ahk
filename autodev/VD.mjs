/** Virtual desktops in the AutoDev language (JavaScript/QuickJS).
 * No AutoHotkey, custom native DLL, Node.js or external COM bridge.
 * Private COM vtables are invoked through the Windows-provided oleaut32.dll.
 */
import {DesktopCom} from './VDCom.mjs';
const integer=(value,label)=>{if(!Number.isInteger(value))throw new RangeError(`${label} must be an integer`);return value;};
export class VirtualDesktops {
 constructor(){
  this.com=new DesktopCom();
  this.user32=ffi.open('user32.dll');
  this.kernel=ffi.open('kernel32.dll');
  this.foreground=this.user32.bind('GetForegroundWindow',{returns:'pointer'});
  this.activate=this.user32.bind('SetForegroundWindow',{args:['pointer'],returns:'bool'});
  this.sleep=this.kernel.bind('Sleep',{args:['u32'],returns:'void'});
  this.closed=false;
 }
 ensure(){if(this.closed)throw Error('VirtualDesktops instance already closed');}
 getCount(){this.ensure();return this.com.count();}
 getCurrentDesktopNum(){this.ensure();return this.com.current();}
 getVersion(){this.ensure();return {build:this.com.build,revision:this.com.revision,abi:this.com.abi};}
 valid(n){integer(n,'Desktop index');const count=this.getCount();if(n<1||n>count)throw RangeError(`Desktop ${n} outside 1..${count}`);return n;}
 modulusResolveDesktopNum(n){integer(n,'Desktop index');const count=this.getCount();if(count<1)throw Error('No virtual desktops');return ((n-1)%count+count)%count+1;}
 waitDesktopSwitched(index,timeoutMs=4000){
  this.valid(index);
  if(!Number.isFinite(timeoutMs)||timeoutMs<0||timeoutMs>120000)throw RangeError('Invalid timeoutMs');
  const started=Date.now();
  do{if(this.getCurrentDesktopNum()===index)return true;if(Date.now()-started>=timeoutMs)return false;this.sleep(25);}while(true);
 }
 goToDesktopNum(index,{wait=true,timeoutMs=4000}={}){
  this.valid(index);
  if(this.getCurrentDesktopNum()!==index)this.com.switch(index);
  if(wait&&!this.waitDesktopSwitched(index,timeoutMs))throw Error(`Timed out switching to desktop ${index}`);
  return index;
 }
 goToRelativeDesktopNum(offset,options){integer(offset,'Desktop offset');return this.goToDesktopNum(this.modulusResolveDesktopNum(this.getCurrentDesktopNum()+offset),options);}
 createDesktop({switchTo=false}={}){
  const count=this.getCount();
  const next=this.com.add();
  if(next<=count)throw Error('Desktop creation did not increase desktop count');
  if(switchTo)this.goToDesktopNum(next);
  return next;
 }
 createUntil(count,{switchToLast=false}={}){
  integer(count,'Desired desktop count');if(count<1||count>100)throw RangeError('Desktop count outside 1..100');
  let n=this.getCount();
  while(n<count)n=this.createDesktop();
  if(switchToLast&&n===count)this.goToDesktopNum(n);
  return n;
 }
 removeDesktop(target=this.getCurrentDesktopNum(),fallback=0){
  this.valid(target);if(this.getCount()===1)throw Error('Cannot remove last desktop');
  if(fallback){this.valid(fallback);if(fallback===target)throw RangeError('Fallback must differ from target');}
  this.com.remove(target,fallback);return this.getCount();
 }
 resolveWindow(value){
  if(typeof value==='bigint'&&value>0n)return value;
  if(typeof value==='number'&&Number.isSafeInteger(value)&&value>0)return BigInt(value);
  if(value&&typeof value==='object'&&value.hwnd)return this.resolveWindow(value.hwnd);
  if(value==='A'){
   const hwnd=this.foreground();if(!hwnd||hwnd===0n)throw Error('No foreground window');return hwnd;
  }
  if(typeof value==='string'&&value){
   const rows=autodev.windows.list().filter(w=>w.hwnd&&(w.title===value||w.title?.includes(value)));
   if(rows.length!==1)throw Error(rows.length?`Ambiguous window title: ${value}`:`Window not found: ${value}`);
   return BigInt(rows[0].hwnd);
  }
  throw TypeError('Specify HWND, {hwnd}, unique window title, or A');
 }
 getDesktopNumOfWindow(target){return this.com.windowDesktop(this.resolveWindow(target));}
 moveWindowToDesktopNum(target,index,{follow=false}={}){
  this.valid(index);const hwnd=this.resolveWindow(target);this.com.moveWindow(hwnd,index);
  if(follow){this.goToDesktopNum(index);this.activate(hwnd);}
  return index;
 }
 moveWindowToRelativeDesktopNum(target,offset,options){integer(offset,'Desktop offset');return this.moveWindowToDesktopNum(target,this.modulusResolveDesktopNum(this.getCurrentDesktopNum()+offset),options);}
 moveWindowToCurrentDesktop(target,options){return this.moveWindowToDesktopNum(target,this.getCurrentDesktopNum(),options);}
 goToDesktopOfWindow(target,{activate=true}={}){
  const hwnd=this.resolveWindow(target),desktop=this.getDesktopNumOfWindow(hwnd);
  if(desktop>0)this.goToDesktopNum(desktop);
  if(activate)this.activate(hwnd);
  return desktop;
 }
 isWindowPinned(target){return this.com.isWindowPinned(this.resolveWindow(target));}
 pinWindow(target){this.com.pinWindow(this.resolveWindow(target),true);}
 unpinWindow(target){this.com.pinWindow(this.resolveWindow(target),false);}
 togglePinWindow(target){this.isWindowPinned(target)?this.unpinWindow(target):this.pinWindow(target);}
 isExePinned(appId){return this.com.isExePinned(appId);}
 pinExe(appId){this.com.pinExe(appId,true);}
 unpinExe(appId){this.com.pinExe(appId,false);}
 togglePinExe(appId){this.isExePinned(appId)?this.unpinExe(appId):this.pinExe(appId);}
 getNameFromDesktopNum(index){this.valid(index);return this.com.getText(index,0);}
 setNameToDesktopNum(index,name){this.valid(index);this.com.setText(index,0,name);}
 getWallpaperFromDesktopNum(index){this.valid(index);return this.com.getText(index,1);}
 setWallpaperToDesktopNum(index,path){this.valid(index);this.com.setText(index,1,path);}
 getInfo(){const current=this.getCurrentDesktopNum(),count=this.getCount(),desktops=[];
  for(let number=1;number<=count;number++){
   let name=null;try{name=this.getNameFromDesktopNum(number);}catch(_ignored){}
   desktops.push({number,name,isCurrent:number===current});
  }
  return {version:this.getVersion(),current,count,desktops};
 }
 close(){if(this.closed)return;this.closed=true;this.user32.close();this.kernel.close();this.com.close();}
}
export const VD=VirtualDesktops;
