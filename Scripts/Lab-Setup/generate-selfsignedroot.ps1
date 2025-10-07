$params = @{
    Type = 'Custom'
    Subject = 'CN=Roanoke Tech Hub Root,OU=www.roanoketechhub.com,O=Roanoke Tech Hub,C=US'
    KeySpec = 'Signature'
    KeyExportPolicy = 'Exportable'
    KeyUsage = 'CertSign'
    KeyUsageProperty = 'Sign'
    KeyLength = 2048
    HashAlgorithm = 'sha256'
    NotAfter = (Get-Date).AddMonths(60)
    CertStoreLocation = 'Cert:\CurrentUser\My'
}
$cert = New-SelfSignedCertificate @params

$params = @{
    Type = 'Custom'
    Subject = 'CN=P2SClient5,OU=www.roanoketechhub.com,O=Roanoke Tech Hub,C=US'
    DnsName = 'P2SChildCert'
    KeySpec = 'Signature'
    KeyExportPolicy = 'Exportable'
    KeyLength = 2048
    HashAlgorithm = 'sha256'
    NotAfter = (Get-Date).AddMonths(18)
    CertStoreLocation = 'Cert:\CurrentUser\My'
    Signer = $cert
    TextExtension = @(
     '2.5.29.37={text}1.3.6.1.5.5.7.3.2')
}
New-SelfSignedCertificate @params