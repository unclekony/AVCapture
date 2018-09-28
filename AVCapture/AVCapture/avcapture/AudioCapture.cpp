#include "AudioCapture.h"
#include <algorithm>


//============================= AudioCaptureGrabberCB

AudioCaptureGrabberCB::AudioCaptureGrabberCB(IAudioCaptureSink *sink)
{
	m_sink = sink;
}

AudioCaptureGrabberCB::~AudioCaptureGrabberCB()
{
	
}

STDMETHODIMP_(ULONG) AudioCaptureGrabberCB::AddRef()
{
	return 2;
}

STDMETHODIMP_(ULONG) AudioCaptureGrabberCB::Release()
{
	return 1;
}

STDMETHODIMP AudioCaptureGrabberCB::QueryInterface(REFIID riid, void ** ppv)

{
	if (riid == IID_ISampleGrabberCB || riid == IID_IUnknown)
	{
		*ppv = (void *) static_cast<ISampleGrabberCB*> (this);
		AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
}

HRESULT STDMETHODCALLTYPE AudioCaptureGrabberCB::SampleCB(double SampleTime, IMediaSample *pSample)
{
	return 0;
}

HRESULT STDMETHODCALLTYPE AudioCaptureGrabberCB::BufferCB(double SampleTime, BYTE *pBuffer, long BufferLen)
{
	if (!pBuffer)
	{
		return E_POINTER;
	}

	if (m_sink)
	{
		m_sink->OnAudioCaptureBufferCB(pBuffer, BufferLen);
	}

	return S_OK;
}

//============================= AudioCapture

AudioCapture::AudioCapture()
{
	CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

	m_graph_builder.CoCreateInstance(CLSID_FilterGraph);
	if (!m_graph_builder)
	{
		OutputDebugString(L"AudioCapture init m_grap_builder failed!");
	}
	else
	{
		OutputDebugString(L"AudioCapture init m_graph_builder succeeded!");
	}
}


AudioCapture::~AudioCapture()
{

	CoUninitialize();
}

HRESULT AudioCapture::initAudioCapture(std::wstring audioDevice, AudioFormat audioFormat)
{
	HRESULT hr;

	m_audio_device = audioDevice;

	hr = initAudioDevice(&m_source);
	if (FAILED(hr))
	{
		OutputDebugString(L"initAudioCapture init audio device fail.");
		return FALSE;
	}

	m_graph_builder->AddFilter(m_source, audioDevice.c_str());
	CComPtr<IPin> source_out_pin = GetOutPin(m_source, 0);

	hr = m_grabber.CoCreateInstance(CLSID_SampleGrabber);
	if (FAILED(hr))
	{
		OutputDebugString(L"initAudioCapture init m_grabber fail.");
		source_out_pin.Release();
		return FALSE;
	}

	AM_MEDIA_TYPE mt;
	ZeroMemory(&mt, sizeof(AM_MEDIA_TYPE));
	mt.majortype = MEDIATYPE_Audio;
	mt.subtype = MEDIASUBTYPE_PCM;
	m_grabber->SetMediaType(&mt);

	CComQIPtr<IBaseFilter, &IID_IBaseFilter> grabber_filter(m_grabber);
	m_graph_builder->AddFilter(grabber_filter, _T("Grabber"));
	IPin *grabber_in_pin = GetInPin(grabber_filter, 0);

	hr = setAudioFormat(audioFormat);
	if (FAILED(hr))
	{
		OutputDebugString(L"initAudioCapture init set audio format fail.");
		m_grabber.Release();
		source_out_pin.Release();
		grabber_filter.Release();
		return FALSE;
	}

	hr = m_graph_builder->Connect(source_out_pin, grabber_in_pin);
	if (FAILED(hr))
	{
		OutputDebugString(L"initAudioCapture init connect pin fail.");
		m_grabber.Release();
		source_out_pin.Release();
		grabber_filter.Release();
		return FALSE;
	}

	// buffering samples as they pass through
	m_grabber->SetBufferSamples(TRUE);

	// not grabbing just one frame
	m_grabber->SetOneShot(FALSE);

	m_grabber_cb = new AudioCaptureGrabberCB(m_sink);
	m_grabber->SetCallback(m_grabber_cb, 1);

	return hr;
}

HRESULT AudioCapture::unInitAudioCapture()
{
	if (m_graph_builder)
	{
		m_graph_builder.Release();
		m_source.Release();
		m_grabber.Release();
		m_grabber_cb.Release();
	}
	return S_OK;
}

HRESULT AudioCapture::listAllAudioDevice(std::list<std::wstring> &deviceList)
{
	CComPtr<IEnumMoniker> enum_moniker;

	CComPtr<ICreateDevEnum> create_dev_enum;
	create_dev_enum.CoCreateInstance(CLSID_SystemDeviceEnum);
	create_dev_enum->CreateClassEnumerator(CLSID_AudioInputDeviceCategory, &enum_moniker, 0);
	create_dev_enum.Release();
	if (!enum_moniker)
	{
		return FALSE;
	}

	enum_moniker->Reset();

	UINT i = 0;
	while (true)
	{
		i++;
		CComPtr<IMoniker> moniker;

		ULONG ulFetched = 0;
		HRESULT hr = enum_moniker->Next(1, &moniker, &ulFetched);
		if (hr != S_OK)
		{
			break;
		}

		CComPtr< IBaseFilter >* ppCap;
		hr = moniker->BindToObject(0, 0, IID_IBaseFilter, (void **)&ppCap);
		if (*ppCap)
		{
			CComPtr< IPropertyBag > pBag;
			hr = moniker->BindToStorage(0, 0, IID_IPropertyBag, (void**)&pBag);
			if (hr != S_OK)
			{
				continue;
			}

			VARIANT var;
			var.vt = VT_BSTR;
			pBag->Read(L"FriendlyName", &var, NULL);


			LPOLESTR wszDeviceID;
			moniker->GetDisplayName(0, NULL, &wszDeviceID);

			deviceList.push_back(var.bstrVal);
		}
	}
	return TRUE;
}

void AudioCapture::setCaptureSink(IAudioCaptureSink *sink)
{
	m_sink = sink;
}

void AudioCapture::startCapture()
{
	CComQIPtr<IMediaControl, &IID_IMediaControl> control(m_graph_builder);
	if (control)
	{
		control->Run();
		control.Release();

		if (m_sink)
		{
			m_sink->OnStartAudioCapture();
		}
	}
}

void AudioCapture::stopCapture()
{
	CComQIPtr<IMediaControl, &IID_IMediaControl> control(m_graph_builder);
	if (control)
	{
		control->Stop();
		control.Release();

		if (m_sink)
		{
			m_sink->OnStopAudioCapture();
		}
	}
}

HRESULT AudioCapture::initAudioDevice(IBaseFilter **pOutFilter)
{
	HRESULT hr = S_OK;

	*pOutFilter = NULL;

	CComPtr<ICreateDevEnum> pCreateDevEnum;
	hr = pCreateDevEnum.CoCreateInstance(CLSID_SystemDeviceEnum);

	if (!pCreateDevEnum || FAILED(hr))
	{
		return hr;
	}

	CComPtr<IEnumMoniker> pEm;
	hr = pCreateDevEnum->CreateClassEnumerator(CLSID_AudioInputDeviceCategory, &pEm, 0);

	if (!pEm || FAILED(hr))
	{
		return hr;
	}

	pEm->Reset();

	ULONG ulFetched = 0;
	CComPtr<IMoniker> pMTarget;
	hr = S_OK;

	for (;;)
	{
		CComPtr<IMoniker> pM;
		hr = pEm->Next(1, &pM, &ulFetched);
		if (hr != S_OK)
		{
			break;
		}

		CComPtr<IPropertyBag> pBag;
		hr = pM->BindToStorage(0, 0, IID_IPropertyBag, (void**)&pBag);
		if (hr != S_OK || !pBag)
		{
			continue;
		}

		VARIANT var;
		var.vt = VT_BSTR;
		hr = pBag->Read(L"FriendlyName", &var, NULL);
		if (FAILED(hr))
		{
			continue;
		}
		std::wstring strDeviceName = var.bstrVal;
		SysFreeString(var.bstrVal);

		if (strDeviceName == m_audio_device)
		{
			pMTarget = pM;
			break;
		}
	}

	if (pMTarget)
	{
		CComPtr<IBaseFilter> pCap;
		hr = pMTarget->BindToObject(0, 0, IID_IBaseFilter, (void**)&pCap);

		*pOutFilter = pCap.Detach();
	}
	else
	{
		OutputDebugString(L"initAudioDevice can not find the target device.");
	}
	return hr;
}

HRESULT AudioCapture::setAudioFormat(AudioFormat audioFormat)
{
	IAMBufferNegotiation *pNeg = NULL;
	WORD wBytesPerSample = audioFormat.m_bits_per_sample / 8;
	DWORD dwBytesPerSecond = wBytesPerSample * audioFormat.m_sample_per_sec * audioFormat.m_channels;

	DWORD dwBufferSize;

	dwBufferSize = 1024 * audioFormat.m_channels * wBytesPerSample;

	// get the nearest, or exact audio format the user wants

	IEnumMediaTypes *pMedia = NULL;
	AM_MEDIA_TYPE *pmt = NULL;

	CComQIPtr<IBaseFilter, &IID_IBaseFilter> cap_filter(m_source);
	CComPtr<IPin> source_output_pin = GetOutPin(cap_filter, 0);
	HRESULT hr = source_output_pin->EnumMediaTypes(&pMedia);

	if (SUCCEEDED(hr))
	{
		while (pMedia->Next(1, &pmt, 0) == S_OK)
		{
			if ((pmt->formattype == FORMAT_WaveFormatEx) &&
				(pmt->cbFormat == sizeof(WAVEFORMATEX)))
			{
				WAVEFORMATEX *wf = (WAVEFORMATEX *)pmt->pbFormat;

				if ((wf->nSamplesPerSec != audioFormat.m_sample_per_sec) ||
					(wf->wBitsPerSample != audioFormat.m_bits_per_sample) ||
					(wf->nChannels != audioFormat.m_channels))
				{
					// found correct audio format
					//
					CComPtr<IAMStreamConfig> pConfig;
					hr = source_output_pin->QueryInterface(IID_IAMStreamConfig, (void **)&pConfig);
					if (SUCCEEDED(hr))
					{
						// get buffer negotiation interface
						source_output_pin->QueryInterface(IID_IAMBufferNegotiation, (void **)&pNeg);

						// set the buffer size based on selected settings
						ALLOCATOR_PROPERTIES prop = { 0 };
						prop.cbBuffer = dwBufferSize;
						prop.cBuffers = 6; // AUDIO STREAM LAG DEPENDS ON THIS
						prop.cbAlign = wBytesPerSample * audioFormat.m_channels;
						pNeg->SuggestAllocatorProperties(&prop);
						SAFE_RELEASE(pNeg);

						WAVEFORMATEX *wf = (WAVEFORMATEX *)pmt->pbFormat;

						// setting additional audio parameters
						wf->nAvgBytesPerSec = dwBytesPerSecond;
						wf->nBlockAlign = wBytesPerSample * audioFormat.m_channels;
						wf->nChannels = audioFormat.m_channels;
						wf->nSamplesPerSec = audioFormat.m_sample_per_sec;
						wf->wBitsPerSample = audioFormat.m_bits_per_sample;

						hr = pConfig->SetFormat(pmt);

					}
					else
					{
						cap_filter.Release();
						source_output_pin.Release();
						pConfig.Release();
						deleteMediaType(pmt);
						SAFE_RELEASE(pMedia);
						// can't set given audio format!
						return FALSE;
					}

					deleteMediaType(pmt);

					hr = pConfig->GetFormat(&pmt);
					if (SUCCEEDED(hr))
					{
						WAVEFORMATEX *wf = (WAVEFORMATEX *)pmt->pbFormat;

						// audio is now initialized
						deleteMediaType(pmt);
						pConfig.Release();
						cap_filter.Release();
						source_output_pin.Release();
						SAFE_RELEASE(pMedia);
						return 0;
					}

					// error initializing audio
					deleteMediaType(pmt);
					pConfig.Release();
					cap_filter.Release();
					source_output_pin.Release();
					SAFE_RELEASE(pMedia);
					return FALSE;
				}
			}
			deleteMediaType(pmt);
		}
		SAFE_RELEASE(pMedia);
	}

	cap_filter.Release();
	source_output_pin.Release();

	return TRUE;
}

IPin* AudioCapture::GetInPin(IBaseFilter* filter, int pinIndex)
{
	return GetPin(filter, PINDIR_INPUT, pinIndex);
}

IPin* AudioCapture::GetOutPin(IBaseFilter* filter, int pinIndex)
{
	return GetPin(filter, PINDIR_OUTPUT, pinIndex);
}

IPin* AudioCapture::GetPin(IBaseFilter* filter, PIN_DIRECTION direction, int index)
{
	IPin* pin = NULL;

	CComPtr< IEnumPins > enum_pins;
	HRESULT hr = filter->EnumPins(&enum_pins);
	if (FAILED(hr))
	{
		return pin;
	}

	while (true)
	{
		ULONG ulFound;
		IPin *tmp_pin;
		hr = enum_pins->Next(1, &tmp_pin, &ulFound);
		if (hr != S_OK)
		{
			break;
		}

		PIN_DIRECTION pindir = (PIN_DIRECTION)3;
		tmp_pin->QueryDirection(&pindir);
		if (pindir == direction)
		{
			if (index == 0)
			{
				tmp_pin->AddRef();
				pin = tmp_pin;
				SAFE_RELEASE(tmp_pin);
				break;
			}
			index--;
		}

		SAFE_RELEASE(tmp_pin);
	}

	return pin;
}

void AudioCapture::deleteMediaType(AM_MEDIA_TYPE *mt)
{
	if (mt == NULL) {
		return;
	}

	//Free MeidaType
	if (mt->cbFormat != 0) {
		CoTaskMemFree((PVOID)mt->pbFormat);

		// Strictly unnecessary but tidier
		mt->cbFormat = 0;
		mt->pbFormat = NULL;
	}
	if (mt->pUnk != NULL) {
		mt->pUnk->Release();
		mt->pUnk = NULL;
	}

	CoTaskMemFree((PVOID)mt);
}