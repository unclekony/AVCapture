#pragma once

#include "../stdafx.h"
#include <dshow.h>
#include <list>
#include <atlbase.h> 

EXTERN_C const CLSID CLSID_SampleGrabber;

#ifndef __ISampleGrabberCB_INTERFACE_DEFINED__
#define __ISampleGrabberCB_INTERFACE_DEFINED__

/* interface ISampleGrabberCB */
/* [unique][helpstring][local][uuid][object] */


EXTERN_C const IID IID_ISampleGrabberCB;


MIDL_INTERFACE("0579154A-2B53-4994-B0D0-E773148EFF85")
ISampleGrabberCB : public IUnknown
{
public:
	virtual HRESULT STDMETHODCALLTYPE SampleCB(
		double SampleTime,
		IMediaSample *pSample) = 0;

	virtual HRESULT STDMETHODCALLTYPE BufferCB(
		double SampleTime,
		BYTE *pBuffer,
		long BufferLen) = 0;

};



#endif 	/* __ISampleGrabberCB_INTERFACE_DEFINED__ */

#ifndef __ISampleGrabber_INTERFACE_DEFINED__
#define __ISampleGrabber_INTERFACE_DEFINED__

/* interface ISampleGrabber */
/* [unique][helpstring][local][uuid][object] */


EXTERN_C const IID IID_ISampleGrabber;


MIDL_INTERFACE("6B652FFF-11FE-4fce-92AD-0266B5D7C78F")
ISampleGrabber : public IUnknown
{
public:
	virtual HRESULT STDMETHODCALLTYPE SetOneShot(
		BOOL OneShot) = 0;

	virtual HRESULT STDMETHODCALLTYPE SetMediaType(
		const AM_MEDIA_TYPE *pType) = 0;

	virtual HRESULT STDMETHODCALLTYPE GetConnectedMediaType(
		AM_MEDIA_TYPE *pType) = 0;

	virtual HRESULT STDMETHODCALLTYPE SetBufferSamples(
		BOOL BufferThem) = 0;

	virtual HRESULT STDMETHODCALLTYPE GetCurrentBuffer(
		/* [out][in] */ long *pBufferSize,
		/* [out] */ long *pBuffer) = 0;

	virtual HRESULT STDMETHODCALLTYPE GetCurrentSample(
		/* [retval][out] */ IMediaSample **ppSample) = 0;

	virtual HRESULT STDMETHODCALLTYPE SetCallback(
		ISampleGrabberCB *pCallback,
		long WhichMethodToCallback) = 0;

};

#endif 	/* __ISampleGrabber_INTERFACE_DEFINED__ */

struct AudioFormat
{
	DWORD m_sample_per_sec;
	WORD m_bits_per_sample;
	WORD m_channels;

	bool operator == (const AudioFormat &right) const
	{
		return m_sample_per_sec == right.m_sample_per_sec && m_bits_per_sample == right.m_bits_per_sample && m_channels == right.m_channels;
	}

	bool operator != (const AudioFormat &right) const {
		return !(operator == (right));
	}

	AudioFormat& operator = (const AudioFormat &right)
	{
		if (this == &right)
		{
			return *this;
		}
		m_sample_per_sec = right.m_sample_per_sec;
		m_bits_per_sample = right.m_bits_per_sample;
		m_channels = right.m_channels;
		return *this;
	}
};

class IAudioCaptureSink
{
public:
	virtual void OnStartAudioCapture() = 0;
	virtual void OnStopAudioCapture() = 0;
	virtual void OnAudioCaptureBufferCB(BYTE *buffer, long bufferSize) = 0;
};


class AudioCaptureGrabberCB : public ISampleGrabberCB
{
public:
	AudioCaptureGrabberCB(IAudioCaptureSink *sink);
	~AudioCaptureGrabberCB();

	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	STDMETHODIMP QueryInterface(REFIID riid, void ** ppv);
	HRESULT STDMETHODCALLTYPE SampleCB(double SampleTime, IMediaSample *pSample);
	HRESULT STDMETHODCALLTYPE BufferCB(double SampleTime, BYTE *pBuffer, long BufferLen);

private:
	IAudioCaptureSink *m_sink;
};

class AudioCapture
{
private:
	AudioCapture();

public:
	static AudioCapture* getInstance(){
		static AudioCapture instance;
		return &instance;
	};

	~AudioCapture();

	HRESULT initAudioCapture(std::wstring audioDevice, AudioFormat audioFormat);
	HRESULT unInitAudioCapture();

	HRESULT listAllAudioDevice(std::list<std::wstring> &deviceList);

	void setCaptureSink(IAudioCaptureSink *sink);

	void startCapture();
	void stopCapture();

	

private:

	HRESULT initAudioDevice(IBaseFilter **pOutFilter);

	HRESULT setAudioFormat(AudioFormat audioFormat);

	IPin* GetInPin(IBaseFilter* filter, int pinIndex);
	IPin* GetOutPin(IBaseFilter* filter, int pinIndex);
	IPin* GetPin(IBaseFilter* filter, PIN_DIRECTION direction, int index);

	void deleteMediaType(AM_MEDIA_TYPE *mt);

	CComPtr<IBaseFilter> m_source;
	CComPtr<IGraphBuilder> m_graph_builder;
	CComPtr<ISampleGrabber> m_grabber;
	CComPtr<AudioCaptureGrabberCB> m_grabber_cb;

	std::wstring m_audio_device;
	AudioFormat m_audio_format;

	IAudioCaptureSink *m_sink;
};
