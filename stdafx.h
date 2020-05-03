#pragma once

#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <SDKDDKVer.h>
#include <windows.h>

#define BMSV_EXPORT	extern "C" __declspec(dllexport)

#include <memory>
#include <atomic>
#include <utility>
