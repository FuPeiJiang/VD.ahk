/* Run ONLY in a disposable KVM guest: exercises real Windows virtual desktops. */
import { VD } from './VD.mjs';
const path=autodev.script.path;
if(!path) throw new Error('Need AutoDev script file execution');
const dir=path.slice(0,Math.max(path.lastIndexOf('\\'),path.lastIndexOf('/'))+1);
const vd=new VD();
const events=[];
function assert(test,label){if(!test)throw new Error('TEST FAILED: '+label);events.push(label);}
let testDesktop=0,original=0,initialCount=0;
try {
  const version=vd.getVersion();
  initialCount=vd.getCount();original=vd.getCurrentDesktopNum();
  assert(version.build>=22000,'version detected');
  assert(initialCount>=1 && original>=1 && original<=initialCount,'initial enumeration');
  assert(vd.modulusResolveDesktopNum(0)===initialCount,'wrap zero');
  assert(vd.waitDesktopSwitched(original,100),'wait at current');
  for(const invalid of [-1,0,initialCount+1]) {
    let rejected=false;try{vd.goToDesktopNum(invalid);}catch(e){rejected=e instanceof RangeError;}
    assert(rejected,`index ${invalid} rejected`);
  }
  testDesktop=vd.createDesktop();
  assert(vd.getCount()===initialCount+1,'created desktop');
  assert(testDesktop===initialCount+1,'returned new desktop index');
  vd.goToDesktopNum(testDesktop);
  assert(vd.waitDesktopSwitched(testDesktop,5000),'switch to created desktop');
  assert(vd.getCurrentDesktopNum()===testDesktop,'current index after switch');
  vd.setNameToDesktopNum(testDesktop,'AutoDev VD KVM smoke');
  assert(vd.getNameFromDesktopNum(testDesktop)==='AutoDev VD KVM smoke','rename and read desktop');
  // AutoDev's own GUI is a real HWND managed through Windows virtual desktops.
  const fixture = gui.window({title:'VD AutoDev disposable KVM fixture',width:420,height:180});
  const hwnd = fixture.hwnd;
  try {
    assert(vd.getDesktopNumOfWindow(hwnd)===testDesktop,'AutoDev GUI belongs to current desktop');
    vd.moveWindowToDesktopNum(hwnd,original);
    assert(vd.getDesktopNumOfWindow(hwnd)===original,'move AutoDev window');
    vd.moveWindowToDesktopNum(hwnd,testDesktop);
    assert(vd.getDesktopNumOfWindow(hwnd)===testDesktop,'move AutoDev window back');
    vd.pinWindow(hwnd);
    assert(vd.isWindowPinned(hwnd),'pin AutoDev window');
    assert(vd.goToDesktopOfWindow(hwnd,{activate:false})<0,'pinned window is not an indexed desktop');
    vd.unpinWindow(hwnd);
    assert(!vd.isWindowPinned(hwnd),'unpin AutoDev window');
  } finally { fixture.close(); }
  vd.goToDesktopNum(original);
  assert(vd.getCurrentDesktopNum()===original,'returned to original desktop');
  vd.removeDesktop(testDesktop,original);
  testDesktop=0;
  assert(vd.getCount()===initialCount,'removed desktop');
  const info=vd.getInfo();
  console.log(JSON.stringify({ok:true,events,info}));
} finally {
  // The KVM overlay is discarded even when a test fails; still clean up on success.
  try { if(original&&vd.getCurrentDesktopNum()!==original)vd.goToDesktopNum(original); }catch(_e){}
  try { if(testDesktop&&vd.getCount()>initialCount)vd.removeDesktop(testDesktop,original||1); }catch(_e){}
  vd.close();
}
