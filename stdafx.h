#pragma once
#define _CRT_SECURE_NO_WARNINGS 1
#define _CRT_SECURE_NO_DEPRECATE 1

#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <SDKDDKVer.h>
#include <windows.h>

#define BMSV_API	__declspec(dllexport)

#include <memory>
#include <atomic>
#include <utility>
