//! デスクトップ音源を指定した期間で録音してファイルに保存するだけのサンプル
//! って思ってたけど、やってみるとマイクから拾った音しか使えていない
//! デスクトップに流れたアプリケーションの音源を記録したい
//! あわよくばアプリケーションを指定して記録したい

use std::{
    fs::File,
    io::BufWriter,
    mem::ManuallyDrop,
    path::PathBuf,
    time::{Duration, Instant},
};

use anyhow::{Context, Result};
use clap::Parser;
use duration_str::parse_std;
use wav::{BitDepth, Header};
use windows::Win32::{
    Media::Audio::{
        eCapture, eConsole, IAudioCaptureClient, IAudioClient, IMMDeviceEnumerator,
        MMDeviceEnumerator, AUDCLNT_SHAREMODE_SHARED, WAVEFORMATEX,
    },
    System::Com::{
        CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_MULTITHREADED,
    },
};

#[derive(Parser, Debug)]
pub struct Cli {
    /// 出力先
    #[clap(short, long)]
    output: PathBuf,

    /// 記録する期間
    #[clap(short, long, default_value = "1m")]
    duration: String,
}

fn main() {
    let cli = Cli::parse();

    // clap の ValueParser 通したいけど今は面倒なのでいい
    let duration = parse_std(&cli.duration).expect("Failed to parse duration text.");

    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED)
            .ok()
            .expect("Failed to initialize COM.")
    }

    let (
        buffer,
        WaveFormatEx {
            format_tag,
            channels,
            samples_per_sec,
            bits_per_sample,
            ..
        },
    ) = unsafe { capture_audio(duration).expect("Failed to capture audio.") };

    let mut output =
        BufWriter::new(File::create(&cli.output).expect("Failed to create output file."));

    let header = Header::new(format_tag, channels, samples_per_sec, bits_per_sample);
    let buffer = BitDepth::Eight(buffer);

    wav::write(header, &buffer, &mut output).expect("Failed to write buffer.");

    unsafe { CoUninitialize() }
}

unsafe fn capture_audio(duration: Duration) -> Result<(Vec<u8>, WaveFormatEx)> {
    let enumerator: IMMDeviceEnumerator = CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
        .context("Failed to create device enumerator.")?;

    let device = enumerator
        .GetDefaultAudioEndpoint(eCapture, eConsole)
        .context("Failed to get default audio endpoint.")?;

    let client: IAudioClient = device
        .Activate(CLSCTX_ALL, None)
        .context("Failed to activate audio client.")?;

    let wave_format = client.GetMixFormat().context("Failed to get mix format.")?;

    let buffered_duration = Duration::from_secs(10);

    client
        .Initialize(
            AUDCLNT_SHAREMODE_SHARED,
            0,
            buffered_duration.as_micros() as i64,
            0,
            wave_format,
            None,
        )
        .context("Failed to initialize audio client.")?;
    let wave_format: WaveFormatEx = (*wave_format).into();

    client.Start().context("Failed to start audio client.")?;

    let cap_client: IAudioCaptureClient = client
        .GetService()
        .context("Failed to get capture client.")?;

    let mut buffer_all = Vec::<u8>::with_capacity(wave_format.avg_bytes_per_sec as usize * 10);
    let started_at = Instant::now();

    while started_at.elapsed() < duration {
        let frames = client
            .GetCurrentPadding()
            .context("Failed to get current padding.")?;
        if frames == 0 {
            continue;
        }

        let mut buffer_ptr: *mut u8 = &mut 0;
        let mut stored_frames = 0;
        let mut flags = 0;
        cap_client.GetBuffer(&mut buffer_ptr, &mut stored_frames, &mut flags, None, None)?;

        let buffer_length = stored_frames * (wave_format.block_align as u32);

        // drop が走っちゃって死ぬので
        let buffer = ManuallyDrop::new(Vec::from_raw_parts(
            buffer_ptr,
            buffer_length as usize,
            buffer_length as usize,
        ));

        buffer_all.extend(buffer.iter());

        cap_client
            .ReleaseBuffer(stored_frames)
            .context("Failed to release buffer.")?;

        std::thread::sleep(Duration::from_micros(100));
    }

    client.Stop().context("Failed to stop client.")?;

    Ok((buffer_all, wave_format))
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
