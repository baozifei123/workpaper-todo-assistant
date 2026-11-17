param(
    [string]$SourceExe = "dist\底稿未完成助手\底稿未完成助手.exe"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$source = Join-Path $root $SourceExe
if (-not (Test-Path $source)) { throw "找不到程序：$source" }
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("wpt_sign_test_" + [guid]::NewGuid())
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
$copy = Join-Path $tempDir "signed_test.exe"
Copy-Item $source $copy
$cert = New-SelfSignedCertificate -Type CodeSigningCert `
    -Subject "CN=WorkpaperTodo Test Only" `
    -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable
try {
    Set-AuthenticodeSignature -FilePath $copy -Certificate $cert `
        -HashAlgorithm SHA256 | Out-Null
    $result = Get-AuthenticodeSignature -FilePath $copy
    Write-Host ("SignatureType=" + $result.SignatureType)
    Write-Host ("Status=" + $result.Status)
    Write-Host ("Signer=" + $result.SignerCertificate.Subject)
    if ($result.SignatureType -ne "Authenticode") {
        throw "签名类型不是 Authenticode"
    }
    Write-Host "自签名流程验证通过；Status=UnknownError 是根证书不受信任的预期结果。"
} finally {
    Remove-Item ("Cert:\CurrentUser\My\" + $cert.Thumbprint) -Force
    Remove-Item -Recurse -Force $tempDir
}
