use std::{collections::VecDeque, mem::ManuallyDrop, ops::Deref, time::Duration};

use anyhow::{ensure, Context as _, Result};
use spectrum_analyzer::{
    samples_fft_to_spectrum, scaling::divide_by_N, windows::hann_window, FrequencyLimit,
};
use windows::Win32::{
    Devices::FunctionDiscovery::PKEY_Device_FriendlyName,
    Media::Audio::{
        eConsole, eRender, IAudioCaptureClient, IAudioClient, IMMDevice, IMMDeviceEnumerator,
        MMDeviceEnumerator, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK, WAVEFORMATEX,
    },
    System::Com::{
        CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_MULTITHREADED,
        STGM_READ,
    },
};

pub fn get_device() -> Result<IMMDevice> {
    unsafe {
        let enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
                .context("Failed to create device enumerator.")?;

        let device = enumerator
            .GetDefaultAudioEndpoint(eRender, eConsole)
            .context("Failed to get default audio endpoint.")?;
        Ok(device)
    }
}

pub fn get_device_name(device: &IMMDevice) -> Result<String> {
    unsafe {
        let store = device.OpenPropertyStore(STGM_READ)?;
        let value = store.GetValue(&PKEY_Device_FriendlyName)?;
        Ok(value.to_string())
    }
}

const SIZE: usize = 2048;

pub struct App {
    name: String,
    client: Client,
    samples: VecDeque<f32>,
    data: Vec<(f64, f64)>,
}

impl App {
    pub fn new(name: String, client: Client) -> App {
        App {
            name,
            client,
            data: Default::default(),
            samples: VecDeque::with_capacity(SIZE),
        }
    }

    pub fn on_tick(&mut self) {
        let format = self.client.wave_format();
        while let Some(buffer) = self.client.get_buffer().expect("Failed to get buffer.") {
            self.samples.extend(
                buffer
                    .chunks(format.block_align as usize)
                    .map(|frame| i16::from_ne_bytes([frame[0], frame[1]]) as f32),
            );
        }
        if self.samples.len() < SIZE {
            return;
        }
        let skips = self.samples.len().saturating_sub(SIZE);
        let samples = self.samples.drain(skips..skips + SIZE).collect::<Vec<_>>();

        let samples = hann_window(&samples);
        let res = samples_fft_to_spectrum(
            &samples,
            format.samples_per_sec,
            FrequencyLimit::Range(60f32, 15_000f32),
            Some(&divide_by_N),
        )
        .unwrap();
        self.data = (1..70)
            .into_iter()
            .map(|freq| freq as f32 * 200.0)
            .map(|freq| (freq as f64, res.freq_val_exact(freq).val().powi(2) as f64))
            .collect();
        // self.data = res
        //     .data()
        //     .into_iter()
        //     .map(|(freq, data)| (freq.val() as f64, (data.val() as f64)))
        //     .collect::<Vec<_>>();
    }

    pub fn data(&self) -> &[(f64, f64)] {
        &self.data
    }
}

pub struct Client {
    device: IMMDevice,
    audio_client: IAudioClient,
    capture_client: IAudioCaptureClient,
    wave_format: WaveFormatEx,
}

impl Client {
    pub fn new(device: IMMDevice) -> Result<Client> {
        unsafe {
            let audio_client: IAudioClient = device
                .Activate(CLSCTX_ALL, None)
                .context("Failed to activate audio client.")?;

            let wave_format = audio_client
                .GetMixFormat()
                .context("Failed to get mix format.")?;

            let buffered_duration = Duration::from_secs(10);

            audio_client
                .Initialize(
                    AUDCLNT_SHAREMODE_SHARED,
                    AUDCLNT_STREAMFLAGS_LOOPBACK,
                    buffered_duration.as_micros() as i64,
                    0,
                    wave_format,
                    None,
                )
                .context("Failed to initialize audio client.")?;
            let wave_format: WaveFormatEx = (*wave_format).into();

            let capture_client: IAudioCaptureClient = audio_client
                .GetService()
                .context("Failed to get capture client.")?;

            audio_client
                .Start()
                .context("Failed to start audio client.")?;
            Ok(Client {
                device,
                audio_client,
                capture_client,
                wave_format,
            })
        }
    }

    pub fn wave_format(&self) -> &WaveFormatEx {
        &self.wave_format
    }

    pub fn get_buffer(&self) -> Result<Option<Vec<u8>>> {
        unsafe {
            let frames = self
                .audio_client
                .GetCurrentPadding()
                .context("Failed to get current padding.")?;
            if frames == 0 {
                return Ok(None);
            }

            let mut buffer_ptr: *mut u8 = &mut 0;
            let mut stored_frames = 0;
            let mut flags = 0;
            self.capture_client
                .GetBuffer(&mut buffer_ptr, &mut stored_frames, &mut flags, None, None)
                .context("Failed to get buffer.")?;

            let buffer_length = stored_frames * (self.wave_format.block_align as u32);
            // drop が走っちゃって死ぬので
            let buffer = ManuallyDrop::new(Vec::from_raw_parts(
                buffer_ptr,
                buffer_length as usize,
                buffer_length as usize,
            ))
            .deref()
            .clone();

            self.capture_client
                .ReleaseBuffer(stored_frames)
                .context("Failed to release buffer.")?;

            Ok(Some(buffer))
        }
    }
}

pub struct Com;

impl Com {
    pub fn initialize() -> Result<Com> {
        unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED)
                .ok()
                .context("Failed to initialize COM.")?;
        }
        Ok(Com)
    }
}

impl Drop for Com {
    fn drop(&mut self) {
        unsafe { CoUninitialize() }
    }
}

#[derive(Debug)]
pub struct WaveFormatEx {
    pub format_tag: u16,
    pub channels: u16,
    pub samples_per_sec: u32,
    pub avg_bytes_per_sec: u32,
    pub block_align: u16,
    pub bits_per_sample: u16,
    pub size: u16,
}

impl From<WAVEFORMATEX> for WaveFormatEx {
    fn from(value: WAVEFORMATEX) -> Self {
        let format_tag = value.wFormatTag;
        let channels = value.nChannels;
        let samples_per_sec = value.nSamplesPerSec;
        let avg_bytes_per_sec = value.nAvgBytesPerSec;
        let block_align = value.nBlockAlign;
        let bits_per_sample = value.wBitsPerSample;
        let size = value.cbSize;
        Self {
            format_tag,
            channels,
            samples_per_sec,
            avg_bytes_per_sec,
            block_align,
            bits_per_sample,
            size,
        }
    }
}
