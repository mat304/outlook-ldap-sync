#include "stdafx.h"
#include "ShimConfig.h"

// Assembly specific values
// TODO: modify the values to reflect those in your project
// assembly strong name
// To Modify: assembly strong name and public key token
//static LPCWSTR szAddInAssemblyName = L"GroupWare, PublicKeyToken=1a9f13915390d0a4";
static LPCWSTR szAddInAssemblyName = L"GroupWare";
// To Modify: full name of connect class
static LPCWSTR szConnectClassName = L"GroupWare.Connect";

// Functions that return the above values

LPCWSTR ShimConfig::AssemblyName()
{
    return szAddInAssemblyName;
}

LPCWSTR ShimConfig::ConnectClassName()
{
    return szConnectClassName;
}
