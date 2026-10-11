/** TypeScript types for VD.mjs on the AutoDev QuickJS runtime. */
export type WindowTarget = number | bigint | {hwnd:number|bigint} | string;
export interface DesktopInfo {
  version: {build:number;revision:number;abi:number};
  count:number;
  current:number;
  desktops:Array<{number:number;name:string|null;isCurrent:boolean}>;
}
export declare class VirtualDesktops {
  constructor();
  getCount():number;
  getCurrentDesktopNum():number;
  getVersion():{build:number;revision:number;abi:number};
  modulusResolveDesktopNum(number:number):number;
  goToDesktopNum(number:number,options?:{wait?:boolean;timeoutMs?:number}):number;
  goToRelativeDesktopNum(offset:number,options?:{wait?:boolean;timeoutMs?:number}):number;
  waitDesktopSwitched(number:number,timeoutMs?:number):boolean;
  createDesktop(options?:{switchTo?:boolean}):number;
  createUntil(count:number,options?:{switchToLast?:boolean}):number;
  removeDesktop(number?:number,fallback?:number):number;
  resolveWindow(target:WindowTarget):bigint;
  getDesktopNumOfWindow(target:WindowTarget):number;
  moveWindowToDesktopNum(target:WindowTarget,number:number,options?:{follow?:boolean}):number;
  moveWindowToRelativeDesktopNum(target:WindowTarget,offset:number,options?:{follow?:boolean}):number;
  moveWindowToCurrentDesktop(target:WindowTarget,options?:{follow?:boolean}):number;
  goToDesktopOfWindow(target:WindowTarget,options?:{activate?:boolean}):number;
  isWindowPinned(target:WindowTarget):boolean;
  pinWindow(target:WindowTarget):void;
  unpinWindow(target:WindowTarget):void;
  togglePinWindow(target:WindowTarget):void;
  isExePinned(appId:string):boolean;
  pinExe(appId:string):void;
  unpinExe(appId:string):void;
  togglePinExe(appId:string):void;
  getNameFromDesktopNum(number:number):string;
  getWallpaperFromDesktopNum(number:number):string;
  setNameToDesktopNum(number:number,name:string):void;
  setWallpaperToDesktopNum(number:number,path:string):void;
  getInfo():DesktopInfo;
  close():void;
}
export {VirtualDesktops as VD};
