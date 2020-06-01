/*
 * (C) Copyright AOE Studio 2010 - All Rights Reserved.
 *
 * This software is the confidential and proprietary information
 * of AOE Studio  ("Confidential Information").  You
 * shall not disclose such Confidential Information and shall use
 * it only in accordance with the terms of the license agreement
 * you entered into with AOE Studio
 *
 */

#include <windows.h>
#include <tchar.h>

BOOL APIENTRY DllMain(HANDLE hModule, 
                      DWORD  ul_reason_for_call, 
                      LPVOID lpReserved
                     )
{
  switch (ul_reason_for_call) {
  case DLL_PROCESS_ATTACH:
    break;

  case DLL_THREAD_ATTACH:
    break;
  case DLL_THREAD_DETACH:
    break;
  case DLL_PROCESS_DETACH:
    break;
  }
  return TRUE;
}