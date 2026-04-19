# Home Migration Guide (GRTC)

## 1) Prepare local repo on office PC

From `C:\Users\user\Desktop\강경현`:

```powershell
& "C:\Program Files\Git\cmd\git.exe" init
& "C:\Program Files\Git\cmd\git.exe" add .
& "C:\Program Files\Git\cmd\git.exe" commit -m "GRTC: upload/parser/simulator updates"
```

If commit fails because of missing identity, set local identity once:

```powershell
& "C:\Program Files\Git\cmd\git.exe" config user.name "GRTC User"
& "C:\Program Files\Git\cmd\git.exe" config user.email "grtc-user@local"
& "C:\Program Files\Git\cmd\git.exe" commit -m "GRTC: upload/parser/simulator updates"
```

## 2) Move to home PC (offline-safe, no cloud required)

Create a single Git bundle file:

```powershell
& "C:\Program Files\Git\cmd\git.exe" bundle create .\grtc-home-transfer.bundle --all
```

Copy only these to USB:

1. `grtc-home-transfer.bundle`
2. `평일 다이아` folder (if needed locally; user already has it at home)

## 3) Restore at home PC

```powershell
mkdir D:\grtc
cd D:\grtc
git clone C:\path\to\grtc-home-transfer.bundle grtc_dial_generator
cd .\grtc_dial_generator
```

## 4) Run at home

```powershell
flutter pub get
flutter run -d chrome
```

## 5) Optional: move changes back to office

At home:

```powershell
git add .
git commit -m "home updates"
git bundle create .\grtc-home-back.bundle --all
```

Bring `grtc-home-back.bundle` to office and import:

```powershell
git fetch C:\path\to\grtc-home-back.bundle "refs/heads/*:refs/remotes/home/*"
git merge home/main
```
