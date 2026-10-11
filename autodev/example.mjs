// autodev run --module example.mjs
import { VD } from './VD.mjs';
const desktops=new VD();
try {
  console.log(JSON.stringify(desktops.getInfo(),null,2));
  // desktops.goToRelativeDesktopNum(1);  // switch right (wraps)
  // desktops.moveWindowToDesktopNum('A', 2, {follow:true});
  // desktops.togglePinWindow('A');
} finally { desktops.close(); }
