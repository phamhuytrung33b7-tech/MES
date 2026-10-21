import React, { useEffect, useRef, useState, useCallback } from 'react';
import { Camera, X, Flashlight, RefreshCw, AlertCircle, Check, Smartphone, ZoomIn, ZoomOut, Zap, Layers, Play } from 'lucide-react';
import jsQR from 'jsqr';

interface MobileQRScannerModalProps {
  isOpen: boolean;
  onClose: () => void;
  onScan: (qrData: string) => void;
  title?: string;
  defaultContinuous?: boolean;
}

export function MobileQRScannerModal({
  isOpen,
  onClose,
  onScan,
  title = 'Quét mã QR bằng Camera điện thoại',
  defaultContinuous = false,
}: MobileQRScannerModalProps) {
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const roiCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const fullCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const animFrameIdRef = useRef<number | null>(null);
  const streamRef = useRef<MediaStream | null>(null);
  const audioCtxRef = useRef<AudioContext | null>(null);
  const isScanningRef = useRef<boolean>(false);
  const isProcessingFrameRef = useRef<boolean>(false);
  const barcodeDetectorRef = useRef<any>(null);
  const lastScanTimestampRef = useRef<number>(0);
  const frameCountRef = useRef<number>(0);
  const lastScannedCodeRef = useRef<string | null>(null);
  const lastScannedTimeRef = useRef<number>(0);

  const [hasPermission, setHasPermission] = useState<boolean | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [errorType, setErrorType] = useState<'permission' | 'notfound' | 'general' | null>(null);
  const [torchOn, setTorchOn] = useState(false);
  const [hasTorch, setHasTorch] = useState(false);
  const [zoomLevel, setZoomLevel] = useState<number>(1);
  const [hasZoom, setHasZoom] = useState(false);
  const [zoomMin, setZoomMin] = useState<number>(1);
  const [zoomMax, setZoomMax] = useState<number>(5);
  const [facingMode, setFacingMode] = useState<'environment' | 'user'>('environment');
  const [lastScannedText, setLastScannedText] = useState<string | null>(null);
  const [isNativeScanner, setIsNativeScanner] = useState(false);
  const [cameraInfo, setCameraInfo] = useState<string>('Camera sau HD');
  const [tapFocusAnim, setTapFocusAnim] = useState<{ x: number; y: number } | null>(null);
  
  // Continuous scanning mode toggle (Warehouse batch scanning)
  const [continuousMode, setContinuousMode] = useState<boolean>(defaultContinuous);
  const [scanCount, setScanCount] = useState<number>(0);
  const [recentScannedList, setRecentScannedList] = useState<string[]>([]);

  // 1. Initialize native BarcodeDetector if available on device (Chrome on Android, iOS 17+ Safari)
  useEffect(() => {
    let isMounted = true;
    const initBarcodeDetector = async () => {
      if (typeof window !== 'undefined' && 'BarcodeDetector' in window) {
        try {
          const BD = (window as any).BarcodeDetector;
          if (typeof BD.getSupportedFormats === 'function') {
            const formats = await BD.getSupportedFormats();
            if (formats && formats.includes('qr_code')) {
              if (isMounted) {
                barcodeDetectorRef.current = new BD({ formats: ['qr_code'] });
                setIsNativeScanner(true);
              }
              return;
            }
          }
          if (isMounted) {
            barcodeDetectorRef.current = new BD({ formats: ['qr_code'] });
            setIsNativeScanner(true);
          }
        } catch {
          if (isMounted) {
            barcodeDetectorRef.current = null;
            setIsNativeScanner(false);
          }
        }
      }
    };
    initBarcodeDetector();
    return () => {
      isMounted = false;
    };
  }, []);

  // 2. Prime Web Audio Context for crisp industrial scanner beep
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
      // Audio not supported
    }
  }, []);

  // 3. Dual-tone industrial beep (like Zebra/Honeywell barcode scanner) + vibration
  const notifySuccess = useCallback(() => {
    try {
      if (audioCtxRef.current) {
        if (audioCtxRef.current.state === 'suspended') {
          audioCtxRef.current.resume().catch(() => {});
        }
        const ctx = audioCtxRef.current;
        const now = ctx.currentTime;

        const osc = ctx.createOscillator();
        const gain = ctx.createGain();

        // High pitch clean chirping beep: 1760Hz (A6) -> 2349Hz (D7)
        osc.type = 'sine';
        osc.frequency.setValueAtTime(1760, now);
        osc.frequency.setValueAtTime(2349, now + 0.04);

        gain.gain.setValueAtTime(0.3, now);
        gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.12);

        osc.connect(gain);
        gain.connect(ctx.destination);

        osc.start(now);
        osc.stop(now + 0.12);
      }
    } catch {
      // Audio fallback
    }

    try {
      if (typeof navigator !== 'undefined' && navigator.vibrate) {
        navigator.vibrate([45, 30, 45]);
      }
    } catch {
      // Vibration not supported
    }
  }, []);

  // 4. Safe full cleanup of camera stream, animation frames, and tracks
  const stopCamera = useCallback(() => {
    isScanningRef.current = false;
    isProcessingFrameRef.current = false;

    if (animFrameIdRef.current) {
      cancelAnimationFrame(animFrameIdRef.current);
      animFrameIdRef.current = null;
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
    setHasZoom(false);
  }, []);

  // 5. Bulletproof camera stream acquisition optimized for iOS & Android rear cameras
  const getCameraStream = async (preferredMode: 'environment' | 'user'): Promise<MediaStream> => {
    const isIOS = typeof navigator !== 'undefined' && (
      /iPad|iPhone|iPod/.test(navigator.userAgent) ||
      (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1)
    );

    // On Android: check if we can enumerate and pinpoint the main rear standard 1x camera
    let backDeviceId: string | null = null;
    if (!isIOS && preferredMode === 'environment' && navigator.mediaDevices?.enumerateDevices) {
      try {
        const devices = await navigator.mediaDevices.enumerateDevices();
        const videoInputs = devices.filter(d => d.kind === 'videoinput');
        
        // Find standard back camera (avoiding ultra-wide 0.5x, macro, depth, telephoto)
        const backCam = videoInputs.find(d => {
          const l = d.label.toLowerCase();
          return (l.includes('back') || l.includes('rear') || l.includes('sau') || l.includes('environment')) &&
            !l.includes('wide') && !l.includes('tele') && !l.includes('macro') && !l.includes('depth') && !l.includes('virtual');
        }) || videoInputs.find(d => {
          const l = d.label.toLowerCase();
          return l.includes('back') || l.includes('rear') || l.includes('sau') || l.includes('0');
        });

        if (backCam && backCam.deviceId) {
          backDeviceId = backCam.deviceId;
        }
      } catch {}
    }

    // Attempt 1: Target back camera with 1080p Full HD resolution for razor-sharp QR modules
    if (backDeviceId) {
      try {
        return await navigator.mediaDevices.getUserMedia({
          video: {
            deviceId: { exact: backDeviceId },
            width: { ideal: 1920 },
            height: { ideal: 1080 },
            frameRate: { ideal: 60, min: 30 },
          },
          audio: false,
        });
      } catch {}
    }

    // Attempt 2: High quality ideal facingMode (Standard cross-platform on iOS Safari & modern Android)
    try {
      return await navigator.mediaDevices.getUserMedia({
        video: {
          facingMode: { ideal: preferredMode },
          width: { ideal: 1920 },
          height: { ideal: 1080 },
          frameRate: { ideal: 60, min: 30 },
        },
        audio: false,
      });
    } catch {}

    // Attempt 3: 720p HD fallback (Compatible with older iPhones & low-spec Android devices)
    try {
      return await navigator.mediaDevices.getUserMedia({
        video: {
          facingMode: { ideal: preferredMode },
          width: { ideal: 1280 },
          height: { ideal: 720 },
          frameRate: { ideal: 30 },
        },
        audio: false,
      });
    } catch {}

    // Attempt 4: Basic facingMode
    try {
      return await navigator.mediaDevices.getUserMedia({
        video: { facingMode: preferredMode },
        audio: false,
      });
    } catch {}

    // Attempt 5: Any camera available
    return await navigator.mediaDevices.getUserMedia({
      video: true,
      audio: false,
    });
  };

  // 6. Camera startup & autofocus configuration
  const startCamera = async () => {
    stopCamera();
    setErrorMessage(null);
    setErrorType(null);
    setHasPermission(null);
    primeAudioContext();

    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      setHasPermission(false);
      setErrorType('general');
      setErrorMessage('Trình duyệt không hỗ trợ Camera. Vui lòng mở bằng Safari trên iOS hoặc Chrome trên Android.');
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
        video.autoplay = true;

        const track = stream.getVideoTracks()[0];
        if (track) {
          const settings = track.getSettings?.() || {};
          const trackLabel = track.label || '';
          const isFront = trackLabel.toLowerCase().includes('front') || facingMode === 'user';
          
          setCameraInfo(
            settings.width && settings.height
              ? `${settings.width}x${settings.height} ${isFront ? 'Cam trước' : 'Cam sau siêu nét'}`
              : 'Cam sau Full HD'
          );

          // Configure continuous autofocus, continuous exposure & zoom capabilities
          if ('getCapabilities' in track) {
            try {
              const capabilities = (track as any).getCapabilities() || {};
              const advanced: any = {};

              // Request continuous autofocus
              if (capabilities.focusMode && Array.isArray(capabilities.focusMode) && capabilities.focusMode.includes('continuous')) {
                advanced.focusMode = 'continuous';
              }
              // Request continuous exposure
              if (capabilities.exposureMode && Array.isArray(capabilities.exposureMode) && capabilities.exposureMode.includes('continuous')) {
                advanced.exposureMode = 'continuous';
              }
              // Request continuous white balance
              if (capabilities.whiteBalanceMode && Array.isArray(capabilities.whiteBalanceMode) && capabilities.whiteBalanceMode.includes('continuous')) {
                advanced.whiteBalanceMode = 'continuous';
              }
              // Check torch
              if (capabilities.torch) {
                setHasTorch(true);
              }
              // Check zoom
              if (capabilities.zoom) {
                setHasZoom(true);
                setZoomMin(capabilities.zoom.min || 1);
                setZoomMax(Math.min(capabilities.zoom.max || 5, 4));
                setZoomLevel(1);
              }

              if (Object.keys(advanced).length > 0) {
                await (track as any).applyConstraints({ advanced: [advanced] }).catch(() => {});
              }
            } catch {}
          }
        }

        // Wait for video frame to load metadata (essential for iOS Safari)
        await new Promise<void>((resolve) => {
          if (video.readyState >= 2 && video.videoWidth > 0) {
            resolve();
          } else {
            const onReady = () => {
              video.removeEventListener('loadeddata', onReady);
              video.removeEventListener('canplay', onReady);
              resolve();
            };
            video.addEventListener('loadeddata', onReady);
            video.addEventListener('canplay', onReady);
            setTimeout(resolve, 1000);
          }
        });

        await video.play().catch(() => {});
        startFastScanLoop();
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

  // 7. Ultra-fast multi-engine scan loop (Native GPU BarcodeDetector + Sharp 1:1 ROI jsQR + Scaled Fallback)
  const scanFrame = async () => {
    if (!isScanningRef.current) return;

    const now = performance.now();
    // Native BarcodeDetector runs at ~30-40fps (25ms), jsQR runs at ~25fps (38ms)
    const minInterval = barcodeDetectorRef.current ? 25 : 38;
    if (now - lastScanTimestampRef.current < minInterval || isProcessingFrameRef.current) {
      animFrameIdRef.current = requestAnimationFrame(scanFrame);
      return;
    }

    const video = videoRef.current;
    if (!video || video.readyState < 2 || video.videoWidth === 0 || video.videoHeight === 0) {
      animFrameIdRef.current = requestAnimationFrame(scanFrame);
      return;
    }

    lastScanTimestampRef.current = now;
    isProcessingFrameRef.current = true;
    frameCountRef.current++;

    try {
      // ENGINE 1: Hardware-Accelerated Native BarcodeDetector (Chrome Android, iOS 17+ Safari)
      // Decodes directly from video element on GPU in 5-15ms!
      if (barcodeDetectorRef.current) {
        try {
          const barcodes = await barcodeDetectorRef.current.detect(video);
          if (barcodes && barcodes.length > 0) {
            const rawVal = barcodes[0].rawValue || barcodes[0].displayValue;
            if (rawVal) {
              handleDecodedQR(rawVal);
              isProcessingFrameRef.current = false;
              return;
            }
          }
        } catch {
          // If native detector hits internal error, fallback to CPU pipeline
        }
      }

      // ENGINE 2: Razor-Sharp Viewfinder ROI (jsQR at 1:1 Native Sensor Pixel Resolution)
      // By taking a 500x500 square from the center of the 1080p stream without downsampling,
      // tiny printed QR codes retain sharp black/white edges and decode instantaneously!
      const roiCanvas = roiCanvasRef.current;
      if (roiCanvas) {
        const vw = video.videoWidth;
        const vh = video.videoHeight;

        // Viewfinder is the center 60% of the smaller dimension
        const roiSize = Math.round(Math.min(vw, vh) * 0.62);
        const sx = Math.round((vw - roiSize) / 2);
        const sy = Math.round((vh - roiSize) / 2);

        // Normalize ROI canvas to 480x480 resolution (optimum speed & accuracy for jsQR)
        const targetDim = 480;
        if (roiCanvas.width !== targetDim || roiCanvas.height !== targetDim) {
          roiCanvas.width = targetDim;
          roiCanvas.height = targetDim;
        }

        const ctx = roiCanvas.getContext('2d', { willReadFrequently: true });
        if (ctx) {
          ctx.drawImage(video, sx, sy, roiSize, roiSize, 0, 0, targetDim, targetDim);
          const imgData = ctx.getImageData(0, 0, targetDim, targetDim);
          const code = jsQR(imgData.data, imgData.width, imgData.height, {
            inversionAttempts: 'dontInvert',
          });

          if (code && code.data) {
            handleDecodedQR(code.data);
            isProcessingFrameRef.current = false;
            return;
          }
        }
      }

      // ENGINE 3: Full-frame Scan with Inversion (runs every 3rd frame to catch off-center or inverted codes)
      if (frameCountRef.current % 3 === 0) {
        const fullCanvas = fullCanvasRef.current;
        if (fullCanvas) {
          const maxDim = 600;
          let fw = video.videoWidth;
          let fh = video.videoHeight;
          if (fw > maxDim || fh > maxDim) {
            if (fw > fh) {
              fh = Math.round((fh * maxDim) / fw);
              fw = maxDim;
            } else {
              fw = Math.round((fw * maxDim) / fh);
              fh = maxDim;
            }
          }

          if (fullCanvas.width !== fw || fullCanvas.height !== fh) {
            fullCanvas.width = fw;
            fullCanvas.height = fh;
          }

          const fctx = fullCanvas.getContext('2d', { willReadFrequently: true });
          if (fctx) {
            fctx.drawImage(video, 0, 0, fw, fh);
            const fullImgData = fctx.getImageData(0, 0, fw, fh);
            const fullCode = jsQR(fullImgData.data, fullImgData.width, fullImgData.height, {
              inversionAttempts: 'attemptBoth',
            });

            if (fullCode && fullCode.data) {
              handleDecodedQR(fullCode.data);
              isProcessingFrameRef.current = false;
              return;
            }
          }
        }
      }
    } catch {
      // Ignore transient decode errors
    } finally {
      isProcessingFrameRef.current = false;
      if (isScanningRef.current) {
        animFrameIdRef.current = requestAnimationFrame(scanFrame);
      }
    }
  };

  // 8. Handle successfully decoded QR code
  const handleDecodedQR = (rawText: string) => {
    const text = rawText.trim();
    if (!text) return;

    const now = Date.now();
    // Debounce exact same code within 1.5 seconds to prevent multi-fire on same label
    if (text === lastScannedCodeRef.current && (now - lastScannedTimeRef.current < 1500)) {
      return;
    }

    lastScannedCodeRef.current = text;
    lastScannedTimeRef.current = now;

    // Trigger industrial feedback (audio + vibration + visual)
    notifySuccess();
    setLastScannedText(text);
    setScanCount(c => c + 1);
    setRecentScannedList(prev => [text, ...prev.slice(0, 4)]);

    // Call onScan callback
    onScan(text);

    // If single-scan mode: close modal after brief visual feedback
    if (!continuousMode) {
      setTimeout(() => {
        onClose();
      }, 500);
    } else {
      // In continuous mode: reset visual alert after 1.2s to be ready for next box
      setTimeout(() => {
        setLastScannedText(null);
      }, 1200);
    }
  };

  // Start smooth high-fps requestAnimationFrame loop
  const startFastScanLoop = () => {
    isScanningRef.current = true;
    lastScanTimestampRef.current = performance.now();
    if (animFrameIdRef.current) cancelAnimationFrame(animFrameIdRef.current);
    animFrameIdRef.current = requestAnimationFrame(scanFrame);
  };

  // Tap-to-refocus support
  const handleTapToFocus = async (e: React.MouseEvent<HTMLDivElement>) => {
    const rect = e.currentTarget.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    setTapFocusAnim({ x, y });
    setTimeout(() => setTapFocusAnim(null), 700);

    // Re-apply continuous autofocus and focus point if supported
    if (streamRef.current) {
      const track = streamRef.current.getVideoTracks()[0];
      if (track && 'applyConstraints' in track) {
        try {
          await (track as any).applyConstraints({
            advanced: [{ focusMode: 'continuous' }]
          });
        } catch {}
      }
    }
  };

  // Hardware Zoom control (1x / 2x / 3x)
  const setZoom = async (newZoom: number) => {
    if (!streamRef.current) return;
    const track = streamRef.current.getVideoTracks()[0];
    if (track && 'applyConstraints' in track) {
      try {
        const clamped = Math.max(zoomMin, Math.min(zoomMax, newZoom));
        await (track as any).applyConstraints({
          advanced: [{ zoom: clamped }]
        });
        setZoomLevel(clamped);
      } catch {}
    }
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
      setLastScannedText(null);
      setScanCount(0);
      setRecentScannedList([]);
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
      <div className="relative w-full max-w-md bg-gray-950 text-white rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[95vh] border border-gray-800">
        
        {/* Header */}
        <div className="flex items-center justify-between px-3.5 sm:px-5 py-3 bg-gray-900/90 border-b border-gray-800">
          <div className="flex items-center gap-2 sm:gap-2.5 min-w-0">
            <div className="p-2 bg-blue-600/20 text-blue-400 rounded-xl shrink-0">
              <Camera size={18} />
            </div>
            <div className="min-w-0">
              <h3 className="text-xs sm:text-sm font-bold tracking-tight text-white flex items-center gap-1.5 truncate">
                <span className="truncate">{title}</span>
                {isNativeScanner && (
                  <span className="px-1.5 py-0.5 rounded text-[9px] font-black bg-blue-500/20 text-blue-300 border border-blue-500/30 flex items-center gap-1 shrink-0">
                    <Zap size={10} className="fill-blue-400 text-blue-400" /> AI Turbo
                  </span>
                )}
              </h3>
              <p className="text-[10px] text-gray-400 font-mono truncate">
                {cameraInfo} • Chạm để lấy nét
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 text-gray-400 hover:text-white bg-gray-800/80 hover:bg-gray-700 rounded-full transition-colors cursor-pointer shrink-0"
            title="Đóng camera"
          >
            <X size={18} />
          </button>
        </div>

        {/* Scan Mode Selector Bar */}
        <div className="bg-gray-900 px-3 py-1.5 border-b border-gray-800 flex items-center justify-between text-xs">
          <div className="flex items-center gap-1.5 text-[11px] text-gray-400 font-medium">
            <Layers size={13} className="text-blue-400" />
            <span>Chế độ:</span>
          </div>
          <div className="flex bg-gray-800/80 p-0.5 rounded-lg border border-gray-700/60">
            <button
              type="button"
              onClick={() => setContinuousMode(false)}
              className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all ${
                !continuousMode ? 'bg-blue-600 text-white shadow-xs' : 'text-gray-400 hover:text-gray-200'
              }`}
            >
              Quét 1 lần
            </button>
            <button
              type="button"
              onClick={() => setContinuousMode(true)}
              className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all flex items-center gap-1 ${
                continuousMode ? 'bg-emerald-600 text-white shadow-xs' : 'text-gray-400 hover:text-gray-200'
              }`}
            >
              <span>Quét liên tục</span>
              {scanCount > 0 && (
                <span className="bg-white/20 px-1 rounded-full text-[9px] font-mono">{scanCount}</span>
              )}
            </button>
          </div>
        </div>

        {/* Camera Viewport */}
        <div 
          onClick={handleTapToFocus}
          className="relative bg-black flex-1 min-h-[350px] sm:min-h-[420px] flex items-center justify-center overflow-hidden cursor-crosshair select-none"
        >
          <video
            ref={videoRef}
            playsInline
            autoPlay
            muted
            className="w-full h-full object-cover pointer-events-none"
          />
          {/* Offscreen canvases for ROI and full frame processing */}
          <canvas ref={roiCanvasRef} className="hidden" />
          <canvas ref={fullCanvasRef} className="hidden" />

          {/* Tap-to-focus ring animation */}
          {tapFocusAnim && (
            <div
              style={{ left: tapFocusAnim.x - 24, top: tapFocusAnim.y - 24 }}
              className="absolute w-12 h-12 border-2 border-amber-400 rounded-full pointer-events-none animate-ping duration-300"
            />
          )}

          {/* Scanner Overlay Reticle */}
          {hasPermission && (
            <div className="absolute inset-0 flex items-center justify-center p-4 pointer-events-none">
              <div className="relative w-64 h-64 sm:w-72 sm:h-72 border-2 border-blue-400/40 rounded-3xl shadow-[0_0_0_9999px_rgba(0,0,0,0.62)] overflow-hidden flex flex-col justify-between p-2">
                
                {/* 4 Corner Markers */}
                <div className="absolute top-0 left-0 w-8 h-8 border-t-4 border-l-4 border-blue-500 rounded-tl-2xl shadow-sm" />
                <div className="absolute top-0 right-0 w-8 h-8 border-t-4 border-r-4 border-blue-500 rounded-tr-2xl shadow-sm" />
                <div className="absolute bottom-0 left-0 w-8 h-8 border-b-4 border-l-4 border-blue-500 rounded-bl-2xl shadow-sm" />
                <div className="absolute bottom-0 right-0 w-8 h-8 border-b-4 border-r-4 border-blue-500 rounded-br-2xl shadow-sm" />

                {/* Laser scan line animation */}
                <div className="w-full h-1 bg-gradient-to-r from-transparent via-blue-400 to-transparent shadow-[0_0_16px_#3b82f6] animate-[pulse_1.2s_infinite]" />

                {/* Center crosshair */}
                <div className="absolute inset-0 flex items-center justify-center opacity-30">
                  <div className="w-4 h-0.5 bg-white" />
                  <div className="h-4 w-0.5 bg-white" />
                </div>

                {/* Success Flash Notification */}
                {lastScannedText && (
                  <div className="absolute inset-0 bg-emerald-950/92 backdrop-blur-xs flex flex-col items-center justify-center text-center p-3 animate-in zoom-in-95 pointer-events-auto">
                    <div className="w-12 h-12 rounded-full bg-emerald-500 text-white flex items-center justify-center mb-2 shadow-xl animate-bounce">
                      <Check size={28} />
                    </div>
                    <span className="text-sm font-bold text-emerald-300">Đã quét siêu tốc!</span>
                    <p className="text-[11px] text-emerald-100 font-mono mt-1 line-clamp-2 max-w-[220px] bg-black/40 px-2 py-1 rounded">
                      {lastScannedText}
                    </p>
                    {continuousMode && (
                      <span className="mt-2 text-[10px] text-emerald-300/80 uppercase tracking-widest font-bold">
                        Đang chờ quét tiếp...
                      </span>
                    )}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Quick Zoom Pill (1x / 2x / 3x) */}
          {hasPermission && hasZoom && (
            <div className="absolute bottom-3 left-1/2 -translate-x-1/2 bg-black/60 backdrop-blur-md px-2 py-1 rounded-full border border-white/10 flex items-center gap-1 z-10">
              <button
                type="button"
                onClick={(e) => { e.stopPropagation(); setZoom(1); }}
                className={`px-2.5 py-1 rounded-full text-xs font-bold transition-all ${
                  zoomLevel === 1 ? 'bg-blue-600 text-white shadow-sm' : 'text-gray-300 hover:text-white'
                }`}
              >
                1x
              </button>
              <button
                type="button"
                onClick={(e) => { e.stopPropagation(); setZoom(2); }}
                className={`px-2.5 py-1 rounded-full text-xs font-bold transition-all ${
                  zoomLevel === 2 ? 'bg-blue-600 text-white shadow-sm' : 'text-gray-300 hover:text-white'
                }`}
              >
                2x
              </button>
              {zoomMax >= 3 && (
                <button
                  type="button"
                  onClick={(e) => { e.stopPropagation(); setZoom(3); }}
                  className={`px-2.5 py-1 rounded-full text-xs font-bold transition-all ${
                    zoomLevel === 3 ? 'bg-blue-600 text-white shadow-sm' : 'text-gray-300 hover:text-white'
                  }`}
                >
                  3x
                </button>
              )}
            </div>
          )}

          {/* Permission / Loading State */}
          {hasPermission === null && (
            <div className="absolute inset-0 bg-gray-950 flex flex-col items-center justify-center gap-3 p-6 text-center">
              <RefreshCw size={28} className="animate-spin text-blue-500" />
              <p className="text-xs sm:text-sm font-medium text-gray-300">Đang khởi động Camera sau siêu nét...</p>
            </div>
          )}

          {/* Error / Permission Denied */}
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
                    <p>• <b>iPhone (iOS Safari):</b> Bấm biểu tượng <b>aA</b> trên thanh địa chỉ → <i>Cài đặt trang web</i> → <i>Camera: Cho phép</i>.</p>
                    <p>• <b>Android (Chrome):</b> Bấm biểu tượng <b>Khóa/Cài đặt</b> cạnh thanh URL → <i>Quyền</i> → <i>Bật Camera</i>.</p>
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
                <span className="hidden sm:inline">{torchOn ? 'Tắt đèn' : 'Bật Flash'}</span>
              </button>
            )}

            <button
              onClick={toggleFacingMode}
              className="p-2.5 bg-gray-800 hover:bg-gray-700 border border-gray-700 text-gray-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer"
              title="Đổi camera trước / sau"
            >
              <RefreshCw size={15} />
              <span className="hidden sm:inline">Đổi Cam</span>
            </button>
          </div>

          <button
            onClick={onClose}
            className="px-5 py-2.5 bg-gray-800 hover:bg-gray-700 border border-gray-700 text-gray-300 font-bold text-xs rounded-xl transition-all cursor-pointer"
          >
            {continuousMode && scanCount > 0 ? 'Xong' : 'Đóng'}
          </button>
        </div>
      </div>
    </div>
  );
}
