#ifndef AUDIOMON_H
#define AUDIOMON_H

#pragma once

BOOL InitVolumeNotification();
void UnregisterVolumeNotifiction();

LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);

static const LPCTSTR AUDIOMON_CLASSNAME = _T("w>audiomon");
static const UINT SNARL_MSG 			= 0x0440;
static const WPARAM SNARL_MSG_MUTED 	= 0;
static const WPARAM SNARL_MSG_VOLUME 	= 1;

// HWND GetSnarlWindow();
HWND GetAudiomonWindow();

#endif