#!/usr/bin/env bash
# Disposable Windows 11 KVM smoke test. NEVER boots/writes the base image.
set -euo pipefail
base=${VD_KVM_BASE:-/var/lib/libvirt/images/autodev-win11-test.qcow2}
vars_base=${VD_KVM_VARS:-/var/lib/libvirt/qemu/nvram/autodev-win11-test_VARS.fd}
bootstrap=${VD_KVM_BOOTSTRAP:-/var/tmp/autodev-kvm-smoke/cache/termirror-4b6d93dd162e2bb852c0.fat.img}
port=${VD_KVM_PORT:-17581}
repo=$(cd "$(dirname "$0")/.." && pwd)
root=$(mktemp -d /var/tmp/vd-kvm-20261011.XXXXXXXX)
qmp=$root/qmp.sock
pid=''
cleanup() {
  if [[ -n $pid ]]; then
    kill "$pid" 2>/dev/null || true
    wait "$pid" 2>/dev/null || true
  fi
  if [[ ${VD_KVM_KEEP:-0} == 1 ]]; then
    echo "KVM test artifacts: $root"
  else
    rm -rf -- "$root"
  fi
}
trap cleanup EXIT
qemu-img create -q -f qcow2 -F qcow2 -b "$base" "$root/overlay.qcow2"
cp "$vars_base" "$root/vars.fd"
install -d -m 0700 "$root/home/.termirror"
install -m 0600 "$HOME/.termirror/identity.key" "$root/home/.termirror/identity.key"
install -m 0644 "$HOME/.termirror/identity.pub" "$root/home/.termirror/identity.pub"
fingerprint=$(sed -n '1p' "$base.termirror-fingerprint" | tr -d '\r\n')
fingerprint=${fingerprint#spki-blake3:}
fingerprint=${fingerprint#spki-blake3 }
printf '127.0.0.1:%s spki-blake3 %s vd-kvm\n' "$port" "$fingerprint" > "$root/home/.termirror/known_servers"
chmod 0600 "$root/home/.termirror/known_servers"
qemu-system-x86_64 -name vd-kvm-20261011 -enable-kvm -machine q35,vmport=off -cpu host \
  -smp 4 -m 6144 \
  -drive if=pflash,format=raw,readonly=on,file=/usr/share/qemu/edk2-x86_64-code.fd \
  -drive "if=pflash,format=raw,file=$root/vars.fd" \
  -drive "file=$root/overlay.qcow2,format=qcow2,if=ide,cache=writeback" \
  -rtc base=localtime,driftfix=slew -boot order=c \
  -netdev "user,id=net0,hostfwd=udp:127.0.0.1:$port-:7443" -device e1000e,netdev=net0 \
  -device qemu-xhci,id=xhci \
  -drive "if=none,id=bootstrap,format=raw,file=$bootstrap,readonly=on" \
  -device usb-storage,drive=bootstrap,bus=xhci.0,removable=on -device usb-tablet,bus=xhci.0 \
  -device qxl-vga -display none -qmp "unix:$qmp,server=on,wait=off" -no-reboot > "$root/qemu.log" 2>&1 &
pid=$!
export HOME="$root/home" TERMIRROR_QUIC_DIRECT=1
target=127.0.0.1:$port
bootstrap_guest() {
python3 - "$qmp" <<'PY'
import json,socket,sys,time
try:
 s=socket.socket(socket.AF_UNIX,socket.SOCK_STREAM);s.settimeout(.4);s.connect(sys.argv[1]);s.recv(65536)
 def send(o):
  s.sendall((json.dumps(o)+'\n').encode());time.sleep(.045)
  try:s.recv(65536)
  except (TimeoutError,socket.timeout):pass
 send({'execute':'qmp_capabilities'})
 def key(k):send({'execute':'human-monitor-command','arguments':{'command-line':'sendkey '+k}})
 key('ret');key('meta_l-r')
 for k in ['d','shift-semicolon','backslash','s','t','a','r','t','minus','t','e','r','m','i','r','r','o','r','dot','c','m','d']:key(k)
 key('ret')
 s.close()
except (OSError,EOFError) as ex: print('KVM QMP bootstrap retry:',ex)
PY
}
echo "Booting disposable KVM guest: $target (overlay=$root/overlay.qcow2)"
ready=0
for i in $(seq 1 110); do
  if timeout -k 1s 2s termirror quic-exec "$target" -- cmd.exe /c echo VD_KVM_READY 2>/dev/null | tr -d '\r' | grep -q '^VD_KVM_READY$'; then
    ready=1; break
  fi
  if (( i % 6 == 0 )); then bootstrap_guest; fi
  sleep 1
done
if (( !ready )); then echo 'KVM guest bootstrap FAILED'; tail -25 "$root/qemu.log"; exit 1; fi
echo 'KVM guest ready; uploading test artifacts'
remote='C:\Users\Administrator\Desktop\VD-AutoDev-KVM'
termirror sync --to "$target" --direct-only --mkpath "$repo/VD.mjs" "$remote\VD.mjs"
termirror sync --to "$target" --direct-only --mkpath "$repo/VDNative.dll" "$remote\VDNative.dll"
termirror sync --to "$target" --direct-only --mkpath "$repo/smoke.mjs" "$remote\smoke.mjs"
termirror sync --to "$target" --direct-only --mkpath /root/GitHub/win-automation-devtools/target/windows-gnu/x86_64-pc-windows-gnu/release/autodev.exe "$remote\autodev.exe"
echo 'AutoDev VD native integration smoke:'
timeout -k 2s 65s termirror quic-exec "$target" -- "$remote\autodev.exe" run --module "$remote\smoke.mjs"
echo 'AUTODEV_VD_KVM_PASSED'
