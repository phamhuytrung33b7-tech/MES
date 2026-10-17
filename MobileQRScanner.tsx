import React, { useEffect, useRef, useState, useCallback } from 'react';
import { Camera, X, Flashlight, RefreshCw, AlertCircle, Check, Smartphone } from 'lucide-react';
import jsQR from 'jsqr';

interface MobileQRScannerModalProps {
  isOpen: boolean;
  onClose: () => void;
  onScan: (qrData: string) => void;
  title?: string;
}

export function MobileQRScannerModal({
  isOpen,
  onClose,
  onScan,
  title = 'Quét mã QR bằng Camera'
}: MobileQRScannerModalProps) {
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const scanTimerRef = useRef<any>(null);
  const streamRef = useRef<MediaStream | null>(null);
  const audioCtxRef = useRef<AudioContext | null>(null);
  const isScanningRef = useRef<boolean>(false);
  const barcodeDetectorRef = useRef<any>(null);

  const [hasPermission, setHasPermission] = useState<boolean | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [errorType, setErrorType] = useState<'permission' | 'notfound' | 'general' | null>(null);
  const [torchOn, setTorchOn] = useState(false);
  const [hasTorch, setHasTorch] = useState(false);
  const [facingMode, setFacingMode] = useState<'environment' | 'user'>('environment');
  const [lastScannedText, setLastScannedText] = useState<string | null>(null);
  const [isNativeScanner, setIsNativeScanner] = useState(false);

  // Initialize native BarcodeDetector if available on device (Chrome on Android, iOS 17+)
  useEffect(() => {
    if (typeof window !== 'undefined' && 'BarcodeDetector' in window) {
      try {
        const BD = (window as any).BarcodeDetector;
        barcodeDetectorRef.current = new BD({ formats: ['qr_code'] });
        setIsNativeScanner(true);
      } catch {
        barcodeDetectorRef.current = null;
        setIsNativeScanner(false);
      }
    }
  }, []);

  // Initialize and prime AudioContext for iOS Safari & Android
  const primeAudioContext = useCallback(() => {
    try {
      if (!audioCtxRef.current) {
        const AudioCtx = window.AudioContext || (window as any).webkitAudioContext;
        if (AudioCtx) {
          audioCtxRef.current = new AudioCtx();
        }
      }
      if (audioCtxRef.current && audioCtxRef.current.state === 'suspended') {
        audioCtxRef.current.resume().catch(() => {});
      }
    } catch {
      // Audio not supported or blocked
    }
  }, []);

  // Play audio beep & vibration on successful scan
  const notifySuccess = useCallback(() => {
    try {
      if (audioCtxRef.current) {
        if (audioCtxRef.current.state === 'suspended') {
          audioCtxRef.current.resume().catch(() => {});
        }
        const ctx = audioCtxRef.current;
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sine';
        osc.frequency.setValueAtTime(880, ctx.currentTime); // 880 Hz (A5)
        gain.gain.setValueAtTime(0.2, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + 0.12);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start();
        osc.stop(ctx.currentTime + 0.12);
      }
    } catch {
      // Audio fallback silent
    }

    try {
      if (typeof navigator !== 'undefined' && navigator.vibrate) {
        navigator.vibrate([80, 40, 80]);
      }
    } catch {
      // Vibration not supported on iOS
    }
  }, []);

  // Full cleanup of camera stream, timers, and tracks
  const stopCamera = useCallback(() => {
    isScanningRef.current = false;
    if (scanTimerRef.current) {
      clearInterval(scanTimerRef.current);
      scanTimerRef.current = null;
    }

    if (streamRef.current) {
      streamRef.current.getTracks().forEach(track => {
        try {
          track.stop();
        } catch {}
      });
      streamRef.current = null;
    }

    if (videoRef.current) {
      videoRef.current.srcObject = null;
    }
    setTorchOn(false);
    setHasTorch(false);
  }, []);

  // Cross-platform camera stream acquisition (iOS Safari & Android optimized fallback ladder)
  const getCameraStream = async (preferredMode: 'environment' | 'user'): Promise<MediaStream> => {
    // Attempt 1: Preferred facingMode with standard mobile 720p resolution
    try {
      return await navigator.mediaDevices.getUserMedia({
        video: {
          facingMode: { ideal: preferredMode },
          width: { ideal: 1280 },
          height: { ideal: 720 },
        },
        audio: false,
      });
    } catch (e1) {
      // Attempt 2: Simple facingMode without resolution constraints (iOS Safari friendly)
      try {
        return await navigator.mediaDevices.getUserMedia({
          video: { facingMode: { ideal: preferredMode } },
          audio: false,
        });
      } catch (e2) {
        // Attempt 3: Any available video device
        return await navigator.mediaDevices.getUserMedia({
          video: true,
          audio: false,
        });
      }
    }
  };

  const startCamera = async () => {
    stopCamera();
    setErrorMessage(null);
    setErrorType(null);
    setHasPermission(null);
    primeAudioContext();

    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      setHasPermission(false);
      setErrorType('general');
      setErrorMessage('Trình duyệt của bạn không hỗ trợ truy cập Camera trực tiếp. Vui lòng sử dụng Safari trên iOS hoặc Chrome trên Android.');
      return;
    }

    try {
      const stream = await getCameraStream(facingMode);
      streamRef.current = stream;
      setHasPermission(true);

      if (videoRef.current) {
        const video = videoRef.current;
        video.srcObject = stream;
        video.setAttribute('playsinline', 'true');
        video.setAttribute('webkit-playsinline', 'true');
        video.muted = true;

        // Check torch capabilities (Android Chrome / supported devices)
        const track = stream.getVideoTracks()[0];
        if (track && 'getCapabilities' in track) {
          try {
            const capabilities = (track as any).getCapabilities();
            if (capabilities && capabilities.torch) {
              setHasTorch(true);
            }
          } catch {}
        }

        // Wait for video to load metadata and begin playing
        await video.play().catch(() => {});
        startScanLoop();
      }
    } catch (err: any) {
      setHasPermission(false);
      const name = err?.name || '';
      if (name === 'NotAllowedError' || name === 'PermissionDeniedError') {
        setErrorType('permission');
        setErrorMessage('Quyền truy cập Camera bị từ chối.');
      } else if (name === 'NotFoundError' || name === 'DevicesNotFoundError') {
        setErrorType('notfound');
        setErrorMessage('Không tìm thấy Camera trên thiết bị.');
      } else {
        setErrorType('general');
        setErrorMessage(err?.message || 'Không thể khởi tạo Camera');
      }
    }
  };

  // Perform single frame detection
  const detectFrame = async () => {
    if (!isScanningRef.current) return;
    const video = videoRef.current;
    if (!video || video.readyState < 2 || video.videoWidth === 0 || video.videoHeight === 0) {
      return;
    }

    try {
      // 1. Hardware-accelerated native BarcodeDetector if available
      if (barcodeDetectorRef.current) {
        const barcodes = await barcodeDetectorRef.current.detect(video);
        if (barcodes && barcodes.length > 0) {
          const rawValue = barcodes[0].rawValue || barcodes[0].displayValue;
          if (rawValue) {
            handleDecodedQR(rawValue);
            return;
          }
        }
      }

      // 2. High-performance jsQR fallback with canvas
      const canvas = canvasRef.current;
      if (!canvas) return;

      // Scale down canvas for faster CPU processing on mobile
      const maxDim = 640;
      let w = video.videoWidth;
      let h = video.videoHeight;
      if (w > maxDim || h > maxDim) {
        if (w > h) {
          h = Math.round((h * maxDim) / w);
          w = maxDim;
        } else {
          w = Math.round((w * maxDim) / h);
          h = maxDim;
        }
      }

      if (canvas.width !== w || canvas.height !== h) {
        canvas.width = w;
        canvas.height = h;
      }

      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      if (ctx) {
        ctx.drawImage(video, 0, 0, w, h);
        const imgData = ctx.getImageData(0, 0, w, h);
        const code = jsQR(imgData.data, imgData.width, imgData.height, {
          inversionAttempts: 'dontInvert',
        });

        if (code && code.data) {
          handleDecodedQR(code.data);
        }
      }
    } catch {
      // Ignore transient decode errors
    }
  };

  const handleDecodedQR = (rawText: string) => {
    const text = rawText.trim();
    if (!text || text === lastScannedText) return;

    setLastScannedText(text);
    notifySuccess();
    onScan(text);

    // Debounce duplicate scans
    setTimeout(() => {
      setLastScannedText(null);
    }, 1600);
  };

  // Throttled scan loop (executes every 120ms = ~8fps, saving 80% battery & CPU on phones)
  const startScanLoop = () => {
    isScanningRef.current = true;
    if (scanTimerRef.current) clearInterval(scanTimerRef.current);
    scanTimerRef.current = setInterval(detectFrame, 120);
  };

  // Flashlight toggle (Android & supported devices)
  const toggleTorch = async () => {
    if (!streamRef.current) return;
    const track = streamRef.current.getVideoTracks()[0];
    if (track && 'applyConstraints' in track) {
      try {
        const nextState = !torchOn;
        await (track as any).applyConstraints({
          advanced: [{ torch: nextState }]
        });
        setTorchOn(nextState);
      } catch {}
    }
  };

  const toggleFacingMode = () => {
    setFacingMode(prev => prev === 'environment' ? 'user' : 'environment');
  };

  useEffect(() => {
    if (isOpen) {
      startCamera();
    } else {
      stopCamera();
    }
    return () => {
      stopCamera();
    };
  }, [isOpen, facingMode]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/85 backdrop-blur-xs p-2 sm:p-6 animate-in fade-in duration-200">
      <div className="relative w-full max-w-md bg-gray-950 text-white rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[94vh] border border-gray-800">
        {/* Header */}
        <div className="flex items-center justify-between px-4 sm:px-5 py-3.5 bg-gray-900/90 border-b border-gray-800">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-blue-600/20 text-blue-400 rounded-xl">
              <Camera size={18} />
            </div>
            <div>
              <h3 className="text-xs sm:text-sm font-bold tracking-tight text-white">{title}</h3>
              <p className="text-[10px] text-gray-400">
                {isNativeScanner ? 'Phần cứng quét AI tăng tốc' : 'Căn chỉnh mã QR vào khung hình'}
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 text-gray-400 hover:text-white bg-gray-800/80 hover:bg-gray-700 rounded-full transition-colors cursor-pointer"
            title="Đóng camera"
          >
            <X size={18} />
          </button>
        </div>

        {/* Camera Viewport */}
        <div className="relative bg-black flex-1 min-h-[320px] sm:min-h-[380px] flex items-center justify-center overflow-hidden">
          <video
            ref={videoRef}
            playsInline
            autoPlay
            muted
            className="w-full h-full object-cover"
          />
          <canvas ref={canvasRef} className="hidden" />

          {/* Scanner Overlay Box */}
          {hasPermission && (
            <div className="absolute inset-0 flex items-center justify-center p-6 pointer-events-none">
              <div className="relative w-64 h-64 border-2 border-blue-400/50 rounded-2xl shadow-[0_0_0_9999px_rgba(0,0,0,0.55)] overflow-hidden flex flex-col justify-between p-2">
                {/* Corner markers */}
                <div className="absolute top-0 left-0 w-8 h-8 border-t-4 border-l-4 border-blue-500 rounded-tl-xl" />
                <div className="absolute top-0 right-0 w-8 h-8 border-t-4 border-r-4 border-blue-500 rounded-tr-xl" />
                <div className="absolute bottom-0 left-0 w-8 h-8 border-b-4 border-l-4 border-blue-500 rounded-bl-xl" />
                <div className="absolute bottom-0 right-0 w-8 h-8 border-b-4 border-r-4 border-blue-500 rounded-br-xl" />

                {/* Laser scan line animation */}
                <div className="w-full h-1 bg-gradient-to-r from-transparent via-blue-400 to-transparent shadow-[0_0_12px_#3b82f6] animate-[pulse_1.5s_infinite]" />

                {lastScannedText && (
                  <div className="absolute inset-0 bg-emerald-950/85 backdrop-blur-xs flex flex-col items-center justify-center text-center p-3 animate-in zoom-in-95 pointer-events-auto">
                    <div className="w-10 h-10 rounded-full bg-emerald-500 text-white flex items-center justify-center mb-2 shadow-lg">
                      <Check size={24} />
                    </div>
                    <span className="text-xs font-bold text-emerald-300">Đã nhận diện thành công!</span>
                    <p className="text-[10px] text-emerald-100 font-mono mt-1 line-clamp-2 max-w-[200px]">
                      {lastScannedText}
                    </p>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Permission / Error Loading State */}
          {hasPermission === null && (
            <div className="absolute inset-0 bg-gray-950 flex flex-col items-center justify-center gap-3 p-6 text-center">
              <RefreshCw size={28} className="animate-spin text-blue-500" />
              <p className="text-xs sm:text-sm font-medium text-gray-300">Đang khởi động Camera...</p>
            </div>
          )}

          {hasPermission === false && (
            <div className="absolute inset-0 bg-gray-950/95 flex flex-col items-center justify-center gap-3.5 p-6 text-center">
              <div className="p-3 bg-red-500/20 text-red-400 rounded-2xl">
                <AlertCircle size={30} />
              </div>
              <div className="space-y-1.5 max-w-xs">
                <p className="text-sm font-bold text-red-300">Chưa cấp quyền Camera</p>
                <p className="text-xs text-gray-400 leading-relaxed">{errorMessage}</p>

                {errorType === 'permission' && (
                  <div className="mt-2 p-2.5 bg-gray-900 border border-gray-800 rounded-xl text-[11px] text-gray-300 text-left space-y-1">
                    <div className="flex items-center gap-1.5 font-bold text-blue-400">
                      <Smartphone size={13} />
                      <span>Cách bật quyền Camera:</span>
                    </div>
                    <p>• <b>iPhone/iOS:</b> Bấm biểu tượng <b>aA</b> trên thanh địa chỉ Safari → <i>Cài đặt trang web</i> → <i>Camera: Cho phép</i>.</p>
                    <p>• <b>Android/Chrome:</b> Bấm biểu tượng <b>Khóa/Cài đặt</b> cạnh thanh URL → <i>Quyền trang web</i> → <i>Bật Camera</i>.</p>
                  </div>
                )}
              </div>

              <button
                onClick={startCamera}
                className="mt-2 px-5 py-2.5 bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs rounded-xl transition-all shadow-md active:scale-95 cursor-pointer"
              >
                Thử lại Camera
              </button>
            </div>
          )}
        </div>

        {/* Footer Actions */}
        <div className="px-4 py-3 bg-gray-900/90 border-t border-gray-800 flex items-center justify-between gap-2.5">
          <div className="flex items-center gap-2">
            {hasTorch && (
              <button
                onClick={toggleTorch}
                className={`p-2.5 rounded-xl border text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer ${
                  torchOn
                    ? 'bg-amber-500/20 text-amber-300 border-amber-500/40'
                    : 'bg-gray-800 text-gray-300 border-gray-700 hover:bg-gray-700'
                }`}
                title="Bật/Tắt Đèn Flash"
              >
                <Flashlight size={15} />
                <span className="hidden sm:inline">{torchOn ? 'Tắt đèn' : 'Đèn Flash'}</span>
              </button>
            )}

            <button
              onClick={toggleFacingMode}
              className="p-2.5 bg-gray-800 hover:bg-gray-700 border border-gray-700 text-gray-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer"
              title="Đổi camera trước / sau"
            >
              <RefreshCw size={15} />
              <span className="hidden sm:inline">Đổi Camera</span>
            </button>
          </div>

          <button
            onClick={onClose}
            className="px-4 py-2.5 bg-gray-800 hover:bg-gray-700 border border-gray-700 text-gray-300 font-bold text-xs rounded-xl transition-all cursor-pointer"
          >
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
}
