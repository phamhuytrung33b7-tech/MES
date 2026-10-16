import React, { useEffect, useRef, useState } from 'react';
import { Camera, X, Flashlight, RefreshCw, QrCode, AlertCircle, Check } from 'lucide-react';
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
  const animFrameIdRef = useRef<number | null>(null);
  const streamRef = useRef<MediaStream | null>(null);

  const [hasPermission, setHasPermission] = useState<boolean | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [torchOn, setTorchOn] = useState(false);
  const [hasTorch, setHasTorch] = useState(false);
  const [facingMode, setFacingMode] = useState<'environment' | 'user'>('environment');
  const [lastScannedText, setLastScannedText] = useState<string | null>(null);

  // Play beep sound on scan success
  const playBeep = () => {
    try {
      const audioCtx = new (window.AudioContext || (window as any).webkitAudioContext)();
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();
      osc.type = 'sine';
      osc.frequency.setValueAtTime(880, audioCtx.currentTime); // A5 note
      gain.gain.setValueAtTime(0.15, audioCtx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.0001, audioCtx.currentTime + 0.15);
      osc.connect(gain);
      gain.connect(audioCtx.destination);
      osc.start();
      osc.stop(audioCtx.currentTime + 0.15);

      if (navigator.vibrate) {
        navigator.vibrate([100, 50, 100]);
      }
    } catch {
      // Audio fallback silent
    }
  };

  const stopCamera = () => {
    if (animFrameIdRef.current) {
      cancelAnimationFrame(animFrameIdRef.current);
      animFrameIdRef.current = null;
    }
    if (streamRef.current) {
      streamRef.current.getTracks().forEach(track => track.stop());
      streamRef.current = null;
    }
  };

  const startCamera = async () => {
    stopCamera();
    setErrorMessage(null);
    setHasPermission(null);

    try {
      const constraints: MediaStreamConstraints = {
        video: {
          facingMode: { ideal: facingMode },
          width: { ideal: 1280 },
          height: { ideal: 720 },
        },
        audio: false,
      };

      const stream = await navigator.mediaDevices.getUserMedia(constraints);
      streamRef.current = stream;
      setHasPermission(true);

      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        videoRef.current.setAttribute('playsinline', 'true');
        await videoRef.current.play();

        // Check if torch capabilities exist
        const track = stream.getVideoTracks()[0];
        if (track && 'getCapabilities' in track) {
          const capabilities = (track as any).getCapabilities();
          if (capabilities.torch) {
            setHasTorch(true);
          }
        }

        scanFrame();
      }
    } catch (err: any) {
      console.error('Camera access error:', err);
      setHasPermission(false);
      if (err.name === 'NotAllowedError' || err.name === 'PermissionDeniedError') {
        setErrorMessage('Quyền truy cập Camera bị từ chối. Vui lòng cấp quyền camera trong cài đặt trình duyệt.');
      } else if (err.name === 'NotFoundError' || err.name === 'DevicesNotFoundError') {
        setErrorMessage('Không tìm thấy Camera trên thiết bị.');
      } else {
        setErrorMessage('Không thể khởi tạo Camera: ' + (err.message || 'Lỗi không xác định'));
      }
    }
  };

  const scanFrame = () => {
    const video = videoRef.current;
    const canvas = canvasRef.current;
    if (!video || !canvas) return;

    if (video.readyState === video.HAVE_ENOUGH_DATA) {
      canvas.width = video.videoWidth;
      canvas.height = video.videoHeight;
      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      if (ctx) {
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const code = jsQR(imageData.data, imageData.width, imageData.height, {
          inversionAttempts: 'dontInvert',
        });

        if (code && code.data) {
          const qrText = code.data.trim();
          if (qrText && qrText !== lastScannedText) {
            setLastScannedText(qrText);
            playBeep();
            onScan(qrText);
            
            // Temporary pause before next scan or auto-close
            setTimeout(() => {
              setLastScannedText(null);
            }, 1500);
          }
        }
      }
    }

    animFrameIdRef.current = requestAnimationFrame(scanFrame);
  };

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
      } catch (e) {
        console.warn('Torch toggle failed:', e);
      }
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
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-xs p-3 sm:p-6 animate-in fade-in duration-200">
      <div className="relative w-full max-w-md bg-gray-900 text-white rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[92vh]">
        {/* Header */}
        <div className="flex items-center justify-between px-5 py-4 bg-gray-900 border-b border-gray-800">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-blue-600/20 text-blue-400 rounded-xl">
              <Camera size={20} />
            </div>
            <div>
              <h3 className="text-sm font-bold tracking-tight text-white">{title}</h3>
              <p className="text-[11px] text-gray-400">Đưa mã QR vào khung hình bên dưới</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-2 text-gray-400 hover:text-white bg-gray-800/80 hover:bg-gray-700 rounded-full transition-colors cursor-pointer"
          >
            <X size={18} />
          </button>
        </div>

        {/* Camera Viewport */}
        <div className="relative bg-black flex-1 min-h-[300px] flex items-center justify-center overflow-hidden">
          <video
            ref={videoRef}
            className="w-full h-full object-cover"
          />
          <canvas ref={canvasRef} className="hidden" />

          {/* Scanner Overlay Box */}
          {hasPermission && (
            <div className="absolute inset-0 flex items-center justify-center p-6">
              <div className="relative w-64 h-64 border-2 border-blue-400/60 rounded-2xl shadow-[0_0_0_9999px_rgba(0,0,0,0.5)] overflow-hidden flex flex-col justify-between p-2">
                {/* Corner markers */}
                <div className="absolute top-0 left-0 w-8 h-8 border-t-4 border-l-4 border-blue-500 rounded-tl-xl" />
                <div className="absolute top-0 right-0 w-8 h-8 border-t-4 border-r-4 border-blue-500 rounded-tr-xl" />
                <div className="absolute bottom-0 left-0 w-8 h-8 border-b-4 border-l-4 border-blue-500 rounded-bl-xl" />
                <div className="absolute bottom-0 right-0 w-8 h-8 border-b-4 border-r-4 border-blue-500 rounded-br-xl" />

                {/* Laser scan line animation */}
                <div className="w-full h-1 bg-gradient-to-r from-transparent via-blue-400 to-transparent shadow-[0_0_12px_#3b82f6] animate-[pulse_1.5s_infinite]" />

                {lastScannedText && (
                  <div className="absolute inset-0 bg-emerald-950/80 backdrop-blur-xs flex flex-col items-center justify-center text-center p-3 animate-in zoom-in-95">
                    <div className="w-10 h-10 rounded-full bg-emerald-500 text-white flex items-center justify-center mb-2 shadow-lg">
                      <Check size={24} />
                    </div>
                    <span className="text-xs font-bold text-emerald-300">Đã nhận dạng thành công!</span>
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
            <div className="absolute inset-0 bg-gray-900 flex flex-col items-center justify-center gap-3 p-6 text-center">
              <RefreshCw size={32} className="animate-spin text-blue-500" />
              <p className="text-sm font-medium text-gray-300">Đang kết nối Camera điện thoại...</p>
            </div>
          )}

          {hasPermission === false && (
            <div className="absolute inset-0 bg-gray-900/95 flex flex-col items-center justify-center gap-4 p-6 text-center">
              <div className="p-3 bg-red-500/20 text-red-400 rounded-2xl">
                <AlertCircle size={32} />
              </div>
              <div className="space-y-1">
                <p className="text-sm font-bold text-red-300">Không thể mở Camera</p>
                <p className="text-xs text-gray-400 max-w-xs">{errorMessage}</p>
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
        <div className="px-5 py-3.5 bg-gray-900 border-t border-gray-800 flex items-center justify-between gap-3">
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
                <Flashlight size={16} />
                <span className="hidden sm:inline">{torchOn ? 'Tắt đèn' : 'Đèn Flash'}</span>
              </button>
            )}

            <button
              onClick={toggleFacingMode}
              className="p-2.5 bg-gray-800 hover:bg-gray-700 border border-gray-700 text-gray-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer"
              title="Đổi camera trước / sau"
            >
              <RefreshCw size={16} />
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
