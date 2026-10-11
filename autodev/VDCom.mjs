// Pure AutoDev COM: Windows system DLLs, no custom native dependencies.
const profiles=[
 [20348,0,'f31574d6-b682-4cdc-bd56-1827860abec6','ff72ffdd-be7e-43fc-9c03-ad81681e88e4',0,4,6,7,9,10,11,-1,-1,-1,-1],
 [22000,0,'094afe11-44f2-4ba0-976f-29a97e263ee0','62fdf88b-11ca-4afb-8bd8-2296dfae49e2',1,4,6,7,9,10,11,6,14,-1,-1],
 [22483,0,'b2f925b9-5a0f-4d2e-9f4d-2b1507593c10','536d3495-b208-4cc9-ae26-de8111275bf8',1,4,6,7,9,10,12,6,15,7,16],
 [22621,2215,'b2f925b9-5a0f-4d2e-9f4d-2b1507593c10','536d3495-b208-4cc9-ae26-de8111275bf8',1,4,6,8,10,11,13,6,16,7,17],
 [22631,3085,'a3175f2d-239c-4bd2-8aa0-eeba8b0b138e','3f07f4be-b107-441a-af0f-39d82529072c',0,4,6,7,9,10,12,5,15,6,16],
 [26100,0,'53f5ca0b-158f-4124-900c-057158060b27','3f07f4be-b107-441a-af0f-39d82529072c',0,4,6,7,9,10,12,5,15,6,16],
 [99999,0,'53f5ca0b-158f-4124-900c-057158060b27','3f07f4be-b107-441a-af0f-39d82529072c',0,4,6,7,9,11,13,5,16,6,17],
];
const octets64=v=>Array.from({length:8},(_,i)=>Number((BigInt(v)>>BigInt(8*i))&255n));
const word64=(b,o=0)=>b.readBytes(o,8).reduceRight((n,x)=>(n<<8n)|BigInt(x),0n);
const assertHr=(v,op)=>{if(v<0)throw Error(op+': HRESULT 0x'+(v>>>0).toString(16));};
export class DesktopCom {
 constructor(){
  this.ole=ffi.open('ole32.dll');this.auto=ffi.open('oleaut32.dll');
  this.nt=ffi.open('ntdll.dll');this.core=ffi.open('combase.dll');this.kernel=ffi.open('kernel32.dll');
  this.pt=ffi.struct({p:'pointer'});
  this.guidParse=this.ole.bind('CLSIDFromString',{args:['wstring','pointer'],returns:'i32'});
  this.create=this.ole.bind('CoCreateInstance',{args:['pointer','pointer','u32','pointer','pointer'],returns:'i32'});
  this.dispatch=this.auto.bind('DispCallFunc',{args:['pointer','usize','i32','u16','u32','pointer','pointer','pointer'],returns:'i32'});
  this.versionlib=ffi.open('version.dll');
  const version=this.readShellVersion();
  this.build=version.build;this.revision=version.revision;
  const i=profiles.findIndex(p=>this.build<p[0]||(this.build===p[0]&&this.revision<p[1]));
  if(i<0)throw Error('Unsupported private desktop ABI');
  this.abi=i;this.p=profiles[i];
 }
 // Match VD.ahk: the private shell ABI follows twinui.pcshell.dll's file version,
 // which can differ from the OSVERSIONINFO build returned by RtlGetVersion.
 readShellVersion(){
  const winDir=ffi.alloc(520),unused=ffi.alloc(4),ptr=ffi.alloc(8),length=ffi.alloc(4);
  let data=null;
  try {
   const systemDir=this.kernel.bind('GetSystemDirectoryW',{args:['pointer','u32'],returns:'u32'});
   const len=systemDir(winDir,260);
   if(!len||len>=260)throw Error('GetSystemDirectoryW failed');
   const chars=winDir.readBytes(0,len*2);let dir='';
   for(let i=0;i<chars.length;i+=2)dir+=String.fromCharCode(chars[i]|(chars[i+1]<<8));
   const target=dir+'\\twinui.pcshell.dll';
   const getSize=this.versionlib.bind('GetFileVersionInfoSizeW',{args:['wstring','pointer'],returns:'u32'});
   const getInfo=this.versionlib.bind('GetFileVersionInfoW',{args:['wstring','u32','u32','pointer'],returns:'bool'});
   const getValue=this.versionlib.bind('VerQueryValueW',{args:['pointer','wstring','pointer','pointer'],returns:'bool'});
   const size=getSize(target,unused);
   if(size<52||size>1048576)throw Error('Unable to read twinui.pcshell.dll version');
   data=ffi.alloc(size);
   if(!getInfo(target,0,size,data)||!getValue(data,'\\',ptr,length))throw Error('Invalid twinui.pcshell.dll version block');
   const relative=word64(ptr)-this.address(data);
   if(relative<0n||relative+16n>BigInt(size)||length.readU32(0)<16)throw Error('Invalid VS_FIXEDFILEINFO pointer');
   const value=data.readU32(Number(relative)+12);
   return {build:value>>>16,revision:value&65535};
  }finally{if(data)data.free();length.free();ptr.free();unused.free();winDir.free();}
 }
 address(b){const s=this.pt.alloc({p:b});try{return s.get('p');}finally{s.free();}}
 guid(s){const b=ffi.alloc(16);try{assertHr(this.guidParse('{'+s+'}',b),'CLSIDFromString');return b;}catch(e){b.free();throw e;}}
 call(ptr,slot,args=[]){
  if(!ptr||slot<0||slot>64)throw Error('Invalid COM interface or vtable slot');
  const buffers=[];const alloc=n=>{const b=ffi.alloc(n);buffers.push(b);return b;};
  try {
   const vt=alloc(Math.max(2,args.length*2)),refs=alloc(Math.max(8,args.length*8)),result=alloc(24),types=[];
   args.forEach((v,i)=>{const cell=alloc(24);cell.writeBytes([21,0]);cell.writeBytes(octets64(v),8);
    refs.writeBytes(octets64(this.address(cell)),i*8);types.push(21,0);});
   vt.writeBytes(types);
   assertHr(this.dispatch(ptr,BigInt(slot*8),4,3,args.length,vt,refs,result),'DispCallFunc');
   assertHr(result.readI32(8),'COM method');
  }finally{while(buffers.length)buffers.pop().free();}
 }
 release(p){if(p)this.call(p,2);}
 out(ptr,slot,args=[],n=8){const b=ffi.alloc(n);try{this.call(ptr,slot,[...args,this.address(b)]);return n===4?b.readU32(0):word64(b);}finally{b.free();}}
 withCom(needView,needPin,fn){
  const guids=[],refs=[];const g=s=>{const b=this.guid(s);guids.push(b);return b;};
  const save=p=>{if(p)refs.push(p);return p;};
  try {
   const b=ffi.alloc(8);let provider=0n;
   try{assertHr(this.create(g('c2f03a33-21f5-47fa-b4bb-156362a2f239'),0n,23,g('6d5140c1-7436-11ce-8034-00aa006009fa'),b),'CoCreateInstance');provider=save(word64(b));}
   finally{b.free();}
   const service=(sid,iid)=>save(this.out(provider,3,[this.address(g(sid)),this.address(g(iid))]));
   const manager=service('c5e0cdca-7b6e-41b2-9fc4-d93975cc467b',this.p[2]);
   const c={manager,view:needView?service('1841c6d7-4f9d-42c0-af41-8747538f10e5','1841c6d7-4f9d-42c0-af41-8747538f10e5'):0n,pin:needPin?service('b5a399e7-1c87-46b8-88e9-fc5747b171bd','4ce81583-1e4c-4632-a621-07a53543148f'):0n};
   return fn(c,save,g);
  }finally{while(refs.length)this.release(refs.pop());while(guids.length)guids.pop().free();}
 }
 list(c,save,g,fn){const ar=save(this.out(c.manager,this.p[7],this.p[4]?[0n]:[]));
  const count=this.out(ar,3,[],4);
  const at=i=>save(this.out(ar,4,[BigInt(i),this.address(g(this.p[3]))]));
  return fn(count,at);
 }
 selected(c,save,g,index,fn){return this.list(c,save,g,(n,at)=>{
  if(!Number.isInteger(index)||index<1||index>n)throw RangeError('Desktop index outside 1..'+n);
  return fn(at(index-1),n);
 });}
 count(){return this.withCom(false,false,(c,s,g)=>this.list(c,s,g,n=>n));}
 current(){return this.withCom(false,false,(c,s,g)=>{
  const now=s(this.out(c.manager,this.p[6],this.p[4]?[0n]:[]));
  const id=ffi.alloc(16);try{this.call(now,4,[this.address(id)]);
   const identity=id.readBytes(0,16);
   return this.list(c,s,g,(n,at)=>{for(let i=0;i<n;i++){
    const guid=ffi.alloc(16);try{this.call(at(i),4,[this.address(guid)]);
     if(guid.readBytes(0,16).every((v,k)=>v===identity[k]))return i+1;
    }finally{guid.free();}}
    throw Error('Current desktop missing from enumeration');
   });
  }finally{id.free();}
 });}
 switch(index){return this.withCom(false,false,(c,s,g)=>this.selected(c,s,g,index,d=>this.call(c.manager,this.p[8],this.p[4]?[0n,d]:[d])));}
 add(){return this.withCom(false,false,(c,s,g)=>{s(this.out(c.manager,this.p[9],this.p[4]?[0n]:[]));return this.list(c,s,g,n=>n);});}
 remove(index,fallback){return this.withCom(false,false,(c,s,g)=>this.list(c,s,g,(n,at)=>{
  if(n<2||index<1||index>n)throw RangeError('Invalid desktop removal');
  if(!fallback)fallback=index===1?2:index-1;
  if(fallback<1||fallback>n||fallback===index)throw RangeError('Invalid fallback');
  this.call(c.manager,this.p[10],[at(index-1),at(fallback-1)]);
 }));}
 viewForWindow(c,hwnd,save){return save(this.out(c.view,6,[BigInt(hwnd)]));}
 moveWindow(hwnd,target){return this.withCom(true,false,(c,s,g)=>this.selected(c,s,g,target,d=>{
  this.call(c.manager,this.p[5],[this.viewForWindow(c,hwnd,s),d]);
 }));}
 isWindowPinned(hwnd){return this.withCom(true,true,(c,s)=>!!this.out(c.pin,6,[this.viewForWindow(c,hwnd,s)],4));}
 pinWindow(hwnd,pin){return this.withCom(true,true,(c,s)=>this.call(c.pin,pin?7:8,[this.viewForWindow(c,hwnd,s)]));}
 windowDesktop(hwnd){return this.withCom(true,false,(c,s,g)=>{
  const view=this.viewForWindow(c,hwnd,s),id=ffi.alloc(16);
  try{
   this.call(view,25,[this.address(id)]);
   const identity=id.readBytes(0,16);
   const matchGuid=(value)=>{const b=g(value);return b.readBytes(0,16).every((v,i)=>v===identity[i]);};
   if(matchGuid('BB64D5B7-4DE3-4AB2-A87C-DB7601AEA7DC'))return -1;
   if(matchGuid('C2DDEA68-66F2-4CF9-8264-1BFD00FBBBAC'))return -2;
   return this.list(c,s,g,(n,at)=>{for(let i=0;i<n;i++){
    const got=ffi.alloc(16);try{this.call(at(i),4,[this.address(got)]);
     if(got.readBytes(0,16).every((v,k)=>v===identity[k]))return i+1;
    }finally{got.free();}}
    throw Error('Window desktop GUID not found');
   });
  }finally{id.free();}
 });}
 withText(value,fn){
  if(typeof value!=='string'||value.length>4096)throw TypeError('Invalid Windows string');
  const b=ffi.alloc((value.length+1)*2);
  try{const chars=Array.from(value+'\0');b.writeBytes(chars.flatMap(ch=>[ch.charCodeAt(0)&255,ch.charCodeAt(0)>>8]));
   return fn(this.address(b));
  }finally{b.free();}
 }
 isExePinned(appId){return this.withCom(false,true,c=>this.withText(appId,p=>!!this.out(c.pin,3,[p],4)));}
 pinExe(appId,pin){return this.withCom(false,true,c=>this.withText(appId,p=>this.call(c.pin,pin?4:5,[p])));}
 getText(index,kind){return this.withCom(false,false,(c,s,g)=>this.selected(c,s,g,index,d=>{
  const slot=this.p[kind?13:11];if(slot<0)throw Error('Metadata unavailable on this Windows build');
  const h=this.out(d,slot),length=ffi.alloc(4);
  try{
   const get=this.core.bind('WindowsGetStringRawBuffer',{args:['pointer','pointer'],returns:'pointer'});
   const ptr=get(h,length),size=length.readU32(0);
   if(size>8192)throw RangeError('Metadata too long');
   if(!size)return '';
   const b=ffi.alloc(size*2);
   try{
    const read=this.kernel.bind('ReadProcessMemory',{args:['pointer','pointer','pointer','usize','pointer'],returns:'bool'});
    const proc=this.kernel.bind('GetCurrentProcess',{returns:'pointer'});
    if(!read(proc(),ptr,b,BigInt(size*2),0n))throw Error('Unable to read metadata');
    const bytes=b.readBytes(0,size*2);let text='';
    for(let i=0;i<bytes.length;i+=2)text+=String.fromCharCode(bytes[i]|bytes[i+1]<<8);
    return text;
   }finally{b.free();}
  }finally{length.free();this.core.bind('WindowsDeleteString',{args:['pointer'],returns:'i32'})(h);}
 }));}
 setText(index,kind,value){return this.withCom(false,false,(c,s,g)=>this.selected(c,s,g,index,d=>{
  const slot=this.p[kind?14:12];if(slot<0)throw Error('Metadata not supported on this Windows build');
  if(typeof value!=='string'||value.length>2048)throw TypeError('Invalid metadata value');
  const h=ffi.alloc(8);
  try {
   const create=this.core.bind('WindowsCreateString',{args:['wstring','u32','pointer'],returns:'i32'});
   const dispose=this.core.bind('WindowsDeleteString',{args:['pointer'],returns:'i32'});
   assertHr(create(value,value.length,h),'WindowsCreateString');
   const handle=word64(h);
   try{this.call(c.manager,slot,[d,handle]);}
   finally{assertHr(dispose(handle),'WindowsDeleteString');}
  }finally{h.free();}
 }));}
 close(){this.versionlib.close();this.kernel.close();this.core.close();this.nt.close();this.auto.close();this.ole.close();}
}
