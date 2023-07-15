function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent();
    $isAdministrator = (New-Object Security.Principal.WindowsPrincipal $currentUser).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

    return $isAdministrator
}
