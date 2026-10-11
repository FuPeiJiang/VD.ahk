// Example AutoDev hotkeys (edit chords to avoid conflicting with Windows).
// Run with: autodev run --module --persistent VD.hotkeys.mjs
import { VD } from './VD.mjs';
const dir=autodev.script.path.replace(/[^\\/]+$/,'');
const vd=new VD({dllPath:dir+'VDNative.dll'});
const report=(operation)=>()=>{
  try { console.log(JSON.stringify({operation,desktop:vd.getCurrentDesktopNum()})); }
  catch(error){ console.log(`VD ${operation} failed: ${String(error)}`); }
};
// Use Alt+Shift combinations by default so Windows' Ctrl+Win combos are retained.
autodev.hotkeys.on('Alt+Shift+Left',()=>{
  try { vd.goToRelativeDesktopNum(-1); report('left')(); } catch(e){ console.log(String(e)); }
});
autodev.hotkeys.on('Alt+Shift+Right',()=>{
  try { vd.goToRelativeDesktopNum(1); report('right')(); } catch(e){ console.log(String(e)); }
});
console.log('VD hotkeys ready: Alt+Shift+Left / Alt+Shift+Right');
