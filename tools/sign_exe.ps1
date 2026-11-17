param(
    [Parameter(Mandatory=$true)][string]$Exe,
    [Parameter(Mandatory=$true)][string]$Pfx,
    [string]$Password = "",
    [string]$Timestamp = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"
if (-not (Test-Path $Exe)) { throw "找不到文件：$Exe" }
if (-not (Test-Path $Pfx)) { throw "找不到证书：$Pfx" }

$signtool = Get-Command signtool.exe -ErrorAction SilentlyContinue
if ($signtool) {
    $signArgs = @("sign", "/fd", "SHA256", "/td", "SHA256", "/tr", $Timestamp,
                  "/f", $Pfx, "/p", $Password, $Exe)
    & $signtool.Source @signArgs
} else {
    Write-Host "未找到 signtool，使用 PowerShell Authenticode 签名（可能不支持时间戳）。"
    $cert = Get-PfxCertificate -FilePath $Pfx
    Set-AuthenticodeSignature -FilePath $Exe -Certificate $cert `
        -HashAlgorithm SHA256 -TimestampServer $Timestamp | Out-Null
}

$result = Get-AuthenticodeSignature -FilePath $Exe
$result | Format-List Status, StatusMessage, SignerCertificate, SignatureType
if ($result.SignatureType -ne "Authenticode") {
    throw "签名验证失败：SignatureType=$($result.SignatureType)"
}
if ($result.Status -ne "Valid") {
    Write-Warning "签名已写入，但证书链状态为 $($result.Status)；正式分发前请确保证书受信任。"
}
