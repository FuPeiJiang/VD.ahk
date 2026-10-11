// autodev run --module example.mjs
import { VD } from './VD.mjs';
const dir=autodev.script.path.replace(/[^\\/]+$/, '');
const desktops=new VD({dllPath:dir+'VDNative.dll'});
try {
  console.log(JSON.stringify(desktops.getInfo(),null,2));
  // desktops.goToRelativeDesktopNum(1);  // switch right (wraps)
  // desktops.moveWindowToDesktopNum('A', 2, {follow:true});
  // desktops.togglePinWindow('A');
} finally { desktops.close(); }
