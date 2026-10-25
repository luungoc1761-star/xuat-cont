import React, { useEffect, useRef, useState, useMemo } from 'react';
import * as THREE from 'three';
import {
  Box,
  RotateCcw,
  Play,
  Pause,
  ZoomIn,
  ZoomOut,
  Layers,
  AlertTriangle,
  CheckCircle2,
  Scale,
  Sparkles,
  Eye
} from 'lucide-react';
import { LeftRightPair, ReconciledItemDemand } from '../types';

interface Warehouse3DSceneProps {
  lrPairs: LeftRightPair[];
  itemDemands: ReconciledItemDemand[];
  selectedModel: string | null;
  onSelectModel: (model: string | null) => void;
  filterMode: 'all_fg' | 'imbalanced' | 'with_demand' | 'cross_wh' | 'all_pairs';
  setFilterMode: (m: 'all_fg' | 'imbalanced' | 'with_demand' | 'cross_wh' | 'all_pairs') => void;
}

export const Warehouse3DScene: React.FC<Warehouse3DSceneProps> = ({
  lrPairs,
  itemDemands,
  selectedModel,
  onSelectModel,
  filterMode,
  setFilterMode
}) => {
  const mountRef = useRef<HTMLDivElement | null>(null);
  const [autoRotate, setAutoRotate] = useState(true);
  const [hoveredPair, setHoveredPair] = useState<LeftRightPair | null>(null);
  const [cameraPreset, setCameraPreset] = useState<'overview' | 'front' | 'top' | 'side'>('overview');

  // Demand lookup by item code
  const demandByCode = useMemo(() => {
    const map = new Map<string, ReconciledItemDemand>();
    for (const d of itemDemands) {
      map.set(d.itemCode.trim(), d);
    }
    return map;
  }, [itemDemands]);

  // Filtered pairs for 3D visualization
  const displayedPairs = useMemo(() => {
    return lrPairs.filter((p) => {
      if (filterMode === 'all_fg') return p.isFinishedGood;
      if (filterMode === 'imbalanced') return !p.isBalanced || p.hasWarehouseCrossBalance;
      if (filterMode === 'cross_wh') return p.leftWh01 > 0 || p.rightWh01 > 0;
      if (filterMode === 'with_demand') {
        return demandByCode.has(p.leftCode) || demandByCode.has(p.rightCode);
      }
      return true;
    });
  }, [lrPairs, filterMode, demandByCode]);

  const activePair = useMemo(() => {
    if (hoveredPair) return hoveredPair;
    if (selectedModel) {
      return lrPairs.find((p) => p.model === selectedModel || p.pairKey === selectedModel) || null;
    }
    return displayedPairs[0] || null;
  }, [hoveredPair, selectedModel, lrPairs, displayedPairs]);

  // Ref synchronization for animation loop
  const autoRotateRef = useRef(autoRotate);
  autoRotateRef.current = autoRotate;
  const selectedModelRef = useRef(selectedModel);
  selectedModelRef.current = selectedModel;

  // Camera spherical coordinates ref
  const cameraStateRef = useRef({
    theta: 0.55,
    phi: 0.98,
    radius: 34,
    targetX: 0,
    targetZ: 0
  });

  useEffect(() => {
    if (cameraPreset === 'overview') {
      cameraStateRef.current = { theta: 0.55, phi: 0.98, radius: 34, targetX: 0, targetZ: 0 };
    } else if (cameraPreset === 'front') {
      cameraStateRef.current = { theta: 0.0, phi: 1.18, radius: 30, targetX: 0, targetZ: 0 };
    } else if (cameraPreset === 'top') {
      cameraStateRef.current = { theta: 0.0, phi: 0.35, radius: 38, targetX: 0, targetZ: 0 };
    } else if (cameraPreset === 'side') {
      cameraStateRef.current = { theta: 1.35, phi: 1.05, radius: 32, targetX: 0, targetZ: 0 };
    }
  }, [cameraPreset]);

  useEffect(() => {
    const container = mountRef.current;
    if (!container) return;

    const width = container.clientWidth || 900;
    const height = container.clientHeight || 480;

    const scene = new THREE.Scene();
    scene.background = new THREE.Color('#04091a');
    scene.fog = new THREE.FogExp2('#04091a', 0.012);

    const camera = new THREE.PerspectiveCamera(45, width / height, 0.1, 250);

    const renderer = new THREE.WebGLRenderer({ antialias: true, powerPreference: 'high-performance' });
    renderer.setSize(width, height);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    renderer.shadowMap.enabled = true;
    renderer.shadowMap.type = THREE.PCFSoftShadowMap;
    container.innerHTML = '';
    container.appendChild(renderer.domElement);

    // Lighting setup
    const ambientLight = new THREE.AmbientLight('#bae6fd', 0.85);
    scene.add(ambientLight);

    const dirLight1 = new THREE.DirectionalLight('#38bdf8', 1.6);
    dirLight1.position.set(22, 35, 24);
    dirLight1.castShadow = true;
    scene.add(dirLight1);

    const dirLight2 = new THREE.DirectionalLight('#10b981', 1.1);
    dirLight2.position.set(-25, 24, -18);
    scene.add(dirLight2);

    const pointLight = new THREE.PointLight('#6366f1', 1.4, 70);
    pointLight.position.set(0, 18, 0);
    scene.add(pointLight);

    // Warehouse Floor Base & Glowing Grid
    const floorGeo = new THREE.PlaneGeometry(76, 56);
    const floorMat = new THREE.MeshStandardMaterial({
      color: '#071026',
      roughness: 0.75,
      metalness: 0.25
    });
    const floorMesh = new THREE.Mesh(floorGeo, floorMat);
    floorMesh.rotation.x = -Math.PI / 2;
    floorMesh.position.y = -0.05;
    floorMesh.receiveShadow = true;
    scene.add(floorMesh);

    const gridHelper = new THREE.GridHelper(72, 36, '#0ea5e9', '#1e293b');
    gridHelper.position.y = 0.01;
    scene.add(gridHelper);

    // Build 3D Twin-Pillar Racks for each Left/Right Pair
    const count = displayedPairs.length;
    const cols = Math.max(4, Math.ceil(Math.sqrt(count * 1.45)));
    const rows = Math.ceil(count / cols);
    const spacingX = 4.2;
    const spacingZ = 3.9;

    const maxSideQty = Math.max(
      1000,
      ...displayedPairs.map((p) => Math.max(p.leftTotal, p.rightTotal))
    );
    const maxBarHeight = 11.5;

    const interactiveMeshes: THREE.Object3D[] = [];
    const pairNodeMap = new Map<string, THREE.Group>();

    // Shared geometries & materials for smooth 60FPS performance
    const padGeo = new THREE.BoxGeometry(3.2, 0.16, 2.5);
    const barGeo = new THREE.BoxGeometry(0.95, 1, 0.95);
    const ringGeo = new THREE.RingGeometry(1.15, 1.38, 28);

    displayedPairs.forEach((pair, idx) => {
      const col = idx % cols;
      const row = Math.floor(idx / cols);
      const x = (col - (cols - 1) / 2) * spacingX;
      const z = (row - (rows - 1) / 2) * spacingZ;

      const group = new THREE.Group();
      group.position.set(x, 0, z);
      group.userData = { pair };

      // Check demand / shortage
      const lDemand = demandByCode.get(pair.leftCode);
      const rDemand = demandByCode.get(pair.rightCode);
      const hasShortage =
        (lDemand && lDemand.shortageQty > 0) || (rDemand && rDemand.shortageQty > 0);

      // Base Pad
      const isSelected = selectedModelRef.current === pair.model;
      const padColor = hasShortage
        ? '#7f1d1d'
        : !pair.isBalanced
        ? '#78350f'
        : pair.hasWarehouseCrossBalance
        ? '#164e63'
        : '#0f172a';
      const padMat = new THREE.MeshStandardMaterial({
        color: padColor,
        emissive: isSelected ? '#0284c7' : hasShortage ? '#dc2626' : !pair.isBalanced ? '#d97706' : '#0f172a',
        emissiveIntensity: isSelected ? 0.65 : hasShortage || !pair.isBalanced ? 0.35 : 0.1,
        metalness: 0.4,
        roughness: 0.4
      });
      const padMesh = new THREE.Mesh(padGeo, padMat);
      padMesh.position.y = 0.08;
      padMesh.userData = { pair };
      group.add(padMesh);
      interactiveMeshes.push(padMesh);

      // Status Halo Ring on Base
      if (!pair.isBalanced || hasShortage || isSelected || pair.hasWarehouseCrossBalance) {
        const ringColor = isSelected
          ? '#38bdf8'
          : hasShortage
          ? '#f43f5e'
          : !pair.isBalanced
          ? '#f59e0b'
          : '#22d3ee';
        const ringMat = new THREE.MeshBasicMaterial({
          color: ringColor,
          side: THREE.DoubleSide,
          transparent: true,
          opacity: 0.75
        });
        const ring = new THREE.Mesh(ringGeo, ringMat);
        ring.rotation.x = -Math.PI / 2;
        ring.position.y = 0.18;
        group.add(ring);
      }

      // Helper to create stacked Left or Right Pillar (Kho 60 bottom + Kho 01 top)
      const createStackedPillar = (
        wh60Qty: number,
        wh01Qty: number,
        offsetX: number,
        color60: string,
        color01: string
      ) => {
        const h60 = wh60Qty > 0 ? Math.max(0.35, (wh60Qty / maxSideQty) * maxBarHeight) : 0;
        const h01 = wh01Qty > 0 ? Math.max(0.35, (wh01Qty / maxSideQty) * maxBarHeight) : 0;

        if (h60 > 0) {
          const mat60 = new THREE.MeshStandardMaterial({
            color: color60,
            emissive: color60,
            emissiveIntensity: 0.22,
            roughness: 0.25,
            metalness: 0.55
          });
          const mesh60 = new THREE.Mesh(barGeo, mat60);
          mesh60.scale.set(1, h60, 1);
          mesh60.position.set(offsetX, 0.16 + h60 / 2, 0);
          mesh60.castShadow = true;
          mesh60.userData = { pair };
          group.add(mesh60);
          interactiveMeshes.push(mesh60);
        }

        if (h01 > 0) {
          const mat01 = new THREE.MeshStandardMaterial({
            color: color01,
            emissive: color01,
            emissiveIntensity: 0.32,
            roughness: 0.2,
            metalness: 0.65
          });
          const mesh01 = new THREE.Mesh(barGeo, mat01);
          mesh01.scale.set(1, h01, 1);
          mesh01.position.set(offsetX, 0.16 + h60 + h01 / 2, 0);
          mesh01.castShadow = true;
          mesh01.userData = { pair };
          group.add(mesh01);
          interactiveMeshes.push(mesh01);
        }

        return h60 + h01;
      };

      // Left Pillar (Trái - LHF): Kho 60 = Cyan (#06b6d4), Kho 01 = Sky/Blue (#3b82f6)
      const leftH = createStackedPillar(pair.leftWh60, pair.leftWh01, -0.68, '#06b6d4', '#3b82f6');
      // Right Pillar (Phải - RHF): Kho 60 = Emerald (#10b981), Kho 01 = Violet (#8b5cf6)
      const rightH = createStackedPillar(pair.rightWh60, pair.rightWh01, 0.68, '#10b981', '#a855f7');

      // Top connector beam showing Left/Right balance or tilt
      const topY = Math.max(leftH, rightH, 0.5) + 0.35;
      const beamGeo = new THREE.BoxGeometry(1.6, 0.12, 0.22);
      const beamMat = new THREE.MeshBasicMaterial({
        color: !pair.isBalanced ? '#f59e0b' : '#38bdf8'
      });
      const beam = new THREE.Mesh(beamGeo, beamMat);
      beam.position.set(0, topY, 0);
      if (!pair.isBalanced) {
        beam.rotation.z = pair.diffLR > 0 ? -0.18 : 0.18;
      }
      group.add(beam);

      // Canvas Sprite Label above each Model Pair
      const labelCanvas = document.createElement('canvas');
      labelCanvas.width = 256;
      labelCanvas.height = 96;
      const ctx = labelCanvas.getContext('2d');
      if (ctx) {
        ctx.fillStyle = isSelected
          ? 'rgba(14, 165, 233, 0.92)'
          : !pair.isBalanced
          ? 'rgba(217, 119, 6, 0.88)'
          : hasShortage
          ? 'rgba(225, 29, 72, 0.88)'
          : 'rgba(15, 23, 42, 0.85)';
        ctx.beginPath();
        ctx.roundRect(8, 8, 240, 80, 14);
        ctx.fill();
        ctx.lineWidth = 3;
        ctx.strokeStyle = !pair.isBalanced ? '#fbbf24' : hasShortage ? '#fb7185' : '#38bdf8';
        ctx.stroke();

        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 28px Inter, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(`#${pair.model}`, 128, 44);

        ctx.fillStyle = '#e2e8f0';
        ctx.font = '600 20px JetBrains Mono, monospace';
        ctx.fillText(`${pair.leftTotal}L | ${pair.rightTotal}R`, 128, 74);
      }
      const tex = new THREE.CanvasTexture(labelCanvas);
      const spriteMat = new THREE.SpriteMaterial({ map: tex, transparent: true });
      const sprite = new THREE.Sprite(spriteMat);
      sprite.position.set(0, topY + 0.95, 0);
      sprite.scale.set(2.5, 0.95, 1);
      group.add(sprite);

      scene.add(group);
      pairNodeMap.set(pair.model, group);
    });

    // Pointer Interaction (Orbit, Pan, Zoom, Raycast Hover & Click)
    const raycaster = new THREE.Raycaster();
    const pointer = new THREE.Vector2();
    let isDragging = false;
    let dragMoved = false;
    let prevX = 0;
    let prevY = 0;

    const onPointerDown = (e: PointerEvent) => {
      isDragging = true;
      dragMoved = false;
      prevX = e.clientX;
      prevY = e.clientY;
    };

    const onPointerMove = (e: PointerEvent) => {
      const rect = renderer.domElement.getBoundingClientRect();
      pointer.x = ((e.clientX - rect.left) / rect.width) * 2 - 1;
      pointer.y = -((e.clientY - rect.top) / rect.height) * 2 + 1;

      if (isDragging) {
        const dx = e.clientX - prevX;
        const dy = e.clientY - prevY;
        if (Math.abs(dx) > 3 || Math.abs(dy) > 3) dragMoved = true;
        prevX = e.clientX;
        prevY = e.clientY;

        cameraStateRef.current.theta -= dx * 0.0075;
        cameraStateRef.current.phi = Math.max(
          0.22,
          Math.min(1.42, cameraStateRef.current.phi - dy * 0.0065)
        );
        return;
      }

      raycaster.setFromCamera(pointer, camera);
      const hits = raycaster.intersectObjects(interactiveMeshes, false);
      if (hits.length > 0 && hits[0].object.userData?.pair) {
        const hitPair = hits[0].object.userData.pair as LeftRightPair;
        setHoveredPair(hitPair);
        renderer.domElement.style.cursor = 'pointer';
      } else {
        setHoveredPair(null);
        renderer.domElement.style.cursor = 'grab';
      }
    };

    const onPointerUp = () => {
      if (isDragging && !dragMoved) {
        raycaster.setFromCamera(pointer, camera);
        const hits = raycaster.intersectObjects(interactiveMeshes, false);
        if (hits.length > 0 && hits[0].object.userData?.pair) {
          const hitPair = hits[0].object.userData.pair as LeftRightPair;
          onSelectModel(
            selectedModelRef.current === hitPair.model ? null : hitPair.model
          );
        }
      }
      isDragging = false;
    };

    const onWheel = (e: WheelEvent) => {
      e.preventDefault();
      cameraStateRef.current.radius = Math.max(
        12,
        Math.min(68, cameraStateRef.current.radius + e.deltaY * 0.025)
      );
    };

    const domEl = renderer.domElement;
    domEl.addEventListener('pointerdown', onPointerDown);
    window.addEventListener('pointermove', onPointerMove);
    window.addEventListener('pointerup', onPointerUp);
    domEl.addEventListener('wheel', onWheel, { passive: false });

    let reqId = 0;
    const animate = () => {
      reqId = requestAnimationFrame(animate);
      if (autoRotateRef.current && !isDragging) {
        cameraStateRef.current.theta += 0.0022;
      }

      const { theta, phi, radius, targetX, targetZ } = cameraStateRef.current;
      camera.position.x = targetX + radius * Math.sin(phi) * Math.sin(theta);
      camera.position.y = radius * Math.cos(phi);
      camera.position.z = targetZ + radius * Math.sin(phi) * Math.cos(theta);
      camera.lookAt(targetX, 2.2, targetZ);

      renderer.render(scene, camera);
    };
    animate();

    const handleResize = () => {
      if (!container) return;
      const w = container.clientWidth || 900;
      const h = container.clientHeight || 480;
      camera.aspect = w / h;
      camera.updateProjectionMatrix();
      renderer.setSize(w, h);
    };
    window.addEventListener('resize', handleResize);

    return () => {
      cancelAnimationFrame(reqId);
      domEl.removeEventListener('pointerdown', onPointerDown);
      window.removeEventListener('pointermove', onPointerMove);
      window.removeEventListener('pointerup', onPointerUp);
      domEl.removeEventListener('wheel', onWheel);
      window.removeEventListener('resize', handleResize);
      renderer.dispose();
    };
  }, [displayedPairs, demandByCode, selectedModel]);

  const activeLeftDemand = activePair ? demandByCode.get(activePair.leftCode) : undefined;
  const activeRightDemand = activePair ? demandByCode.get(activePair.rightCode) : undefined;

  return (
    <section className="rounded-2xl border border-slate-800/90 bg-slate-900/75 backdrop-blur-xl shadow-2xl overflow-hidden">
      {/* Top Header of 3D Section */}
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-800/80 bg-slate-900/90 px-5 py-4">
        <div className="flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-gradient-to-br from-cyan-500 to-emerald-500 text-slate-950 shadow-lg shadow-cyan-500/20">
            <Box className="h-5 w-5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-base sm:text-lg font-bold tracking-tight text-white">
                Báo Cáo Dashboard 3D Động — Tồn Kho Thành Phẩm Kho 01 & Kho 60 (Trái / Phải)
              </h2>
              <span className="inline-flex items-center gap-1 rounded-full bg-cyan-500/15 border border-cyan-500/30 px-2.5 py-0.5 text-xs font-semibold text-cyan-300">
                <Sparkles className="h-3 w-3" /> WebGL Realtime 3D
              </span>
            </div>
            <p className="text-xs text-slate-400">
              Mô phỏng không gian 3D từng mã Model/Feature theo cặp cột đôi Trái (LHF) & Phải (RHF) phân tầng Kho 60 + Kho 01
            </p>
          </div>
        </div>

        {/* Filter Mode Pills */}
        <div className="flex flex-wrap items-center gap-1.5">
          {[
            { id: 'all_fg', label: '51 Cặp Thành Phẩm (Đầu 8)', count: lrPairs.filter((x) => x.isFinishedGood).length },
            {
              id: 'imbalanced',
              label: 'Lệch Cặp / Cân Bằng Chéo Kho',
              count: lrPairs.filter((x) => !x.isBalanced || x.hasWarehouseCrossBalance).length
            },
            {
              id: 'with_demand',
              label: 'Đang Có Đơn Xuất',
              count: lrPairs.filter(
                (x) => demandByCode.has(x.leftCode) || demandByCode.has(x.rightCode)
              ).length
            },
            {
              id: 'cross_wh',
              label: 'Có Tồn Ở Kho 01',
              count: lrPairs.filter((x) => x.leftWh01 > 0 || x.rightWh01 > 0).length
            },
            { id: 'all_pairs', label: 'Tất Cả Cặp L/R', count: lrPairs.length }
          ].map((tab) => (
            <button
              key={tab.id}
              onClick={() => setFilterMode(tab.id as any)}
              className={`inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-semibold transition-all cursor-pointer ${
                filterMode === tab.id
                  ? 'bg-cyan-500 text-slate-950 shadow-md shadow-cyan-500/25'
                  : 'bg-slate-800/80 text-slate-300 hover:bg-slate-800 hover:text-white border border-slate-700/60'
              }`}
            >
              <span>{tab.label}</span>
              <span
                className={`rounded-full px-1.5 py-0.2 text-[10px] font-mono ${
                  filterMode === tab.id ? 'bg-slate-950/25 text-slate-950' : 'bg-slate-900 text-cyan-300'
                }`}
              >
                {tab.count}
              </span>
            </button>
          ))}
        </div>
      </div>

      {/* Main 3D Viewport + Side Inspector Grid */}
      <div className="grid grid-cols-1 lg:grid-cols-12">
        {/* 3D Canvas Area */}
        <div className="relative lg:col-span-8 h-[460px] sm:h-[510px] bg-slate-950 overflow-hidden">
          <div ref={mountRef} className="w-full h-full" />

          {/* Top-Left Legend Overlay */}
          <div className="pointer-events-none absolute top-3 left-3 flex flex-wrap gap-2 rounded-xl bg-slate-900/85 border border-slate-800/90 backdrop-blur-md px-3 py-2 text-[11px] text-slate-200">
            <span className="inline-flex items-center gap-1.5">
              <span className="h-2.5 w-2.5 rounded-xs bg-cyan-500 inline-block" /> Trái (LHF) - Kho 60
            </span>
            <span className="inline-flex items-center gap-1.5">
              <span className="h-2.5 w-2.5 rounded-xs bg-blue-500 inline-block" /> Trái (LHF) - Kho 01
            </span>
            <span className="inline-flex items-center gap-1.5">
              <span className="h-2.5 w-2.5 rounded-xs bg-emerald-500 inline-block" /> Phải (RHF) - Kho 60
            </span>
            <span className="inline-flex items-center gap-1.5">
              <span className="h-2.5 w-2.5 rounded-xs bg-purple-500 inline-block" /> Phải (RHF) - Kho 01
            </span>
            <span className="inline-flex items-center gap-1.5">
              <span className="h-2.5 w-2.5 rounded-full border-2 border-amber-400 inline-block" /> Lệch Cặp L/R
            </span>
            <span className="inline-flex items-center gap-1.5">
              <span className="h-2.5 w-2.5 rounded-full border-2 border-rose-500 inline-block" /> Thiếu Hàng Xuất
            </span>
          </div>

          {/* Bottom-Left Camera Controls */}
          <div className="absolute bottom-3 left-3 flex flex-wrap items-center gap-1.5 rounded-xl bg-slate-900/90 border border-slate-800/90 backdrop-blur-md p-1.5">
            <button
              onClick={() => setAutoRotate(!autoRotate)}
              className={`inline-flex items-center gap-1 rounded-lg px-2.5 py-1.5 text-xs font-medium transition cursor-pointer ${
                autoRotate
                  ? 'bg-cyan-500/20 text-cyan-300 border border-cyan-500/40'
                  : 'bg-slate-800 text-slate-300 hover:text-white'
              }`}
              title="Bật/Tắt tự động xoay 3D"
            >
              {autoRotate ? <Pause className="h-3.5 w-3.5" /> : <Play className="h-3.5 w-3.5" />}
              <span>{autoRotate ? 'Đang Xoay 3D' : 'Tạm Dừng'}</span>
            </button>

            {(['overview', 'front', 'top', 'side'] as const).map((preset) => (
              <button
                key={preset}
                onClick={() => setCameraPreset(preset)}
                className={`rounded-lg px-2.5 py-1.5 text-xs font-medium transition cursor-pointer ${
                  cameraPreset === preset
                    ? 'bg-slate-700 text-white'
                    : 'text-slate-400 hover:bg-slate-800 hover:text-slate-200'
                }`}
              >
                {preset === 'overview'
                  ? 'Góc 3D Chuẩn'
                  : preset === 'front'
                  ? 'Chính Diện L/R'
                  : preset === 'top'
                  ? 'Mặt Bằng Kho'
                  : 'Góc Nghiêng'}
              </button>
            ))}

            <button
              onClick={() => {
                cameraStateRef.current.radius = Math.max(14, cameraStateRef.current.radius - 4);
              }}
              className="rounded-lg p-1.5 text-slate-300 hover:bg-slate-800 hover:text-white cursor-pointer"
              title="Phóng to"
            >
              <ZoomIn className="h-4 w-4" />
            </button>
            <button
              onClick={() => {
                cameraStateRef.current.radius = Math.min(60, cameraStateRef.current.radius + 4);
              }}
              className="rounded-lg p-1.5 text-slate-300 hover:bg-slate-800 hover:text-white cursor-pointer"
              title="Thu nhỏ"
            >
              <ZoomOut className="h-4 w-4" />
            </button>
            <button
              onClick={() => {
                setCameraPreset('overview');
                onSelectModel(null);
              }}
              className="rounded-lg p-1.5 text-slate-300 hover:bg-slate-800 hover:text-white cursor-pointer"
              title="Đặt lại góc nhìn"
            >
              <RotateCcw className="h-4 w-4" />
            </button>
          </div>

          {/* Bottom-Right Help Hint */}
          <div className="pointer-events-none hidden sm:flex absolute bottom-3 right-3 items-center gap-1.5 rounded-lg bg-slate-900/80 border border-slate-800 px-2.5 py-1 text-[11px] text-slate-400">
            <Eye className="h-3.5 w-3.5 text-cyan-400" />
            <span>Kéo chuột để xoay 360° • Cuộn để Phóng to/Thu nhỏ • Click vào tháp để khóa mã hàng</span>
          </div>
        </div>

        {/* Right Side: Live 3D Model Inspector + Left/Right Pair Balance List */}
        <div className="lg:col-span-4 flex flex-col border-t lg:border-t-0 lg:border-l border-slate-800/80 bg-slate-900/95 h-[510px]">
          {/* Active Pair Inspector Card */}
          {activePair && (
            <div className="p-4 border-b border-slate-800/90 bg-gradient-to-b from-slate-900 to-slate-950/80">
              <div className="flex items-start justify-between gap-2">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-mono uppercase tracking-wider text-cyan-400">
                      MODEL / FEATURE #{activePair.model}
                    </span>
                    {activePair.isBalanced ? (
                      <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/15 border border-emerald-500/30 px-2 py-0.5 text-[11px] font-semibold text-emerald-300">
                        <CheckCircle2 className="h-3 w-3" />
                        {activePair.hasWarehouseCrossBalance ? 'Cân Bằng Khi Gộp 01+60' : 'Cân Bằng L/R 100%'}
                      </span>
                    ) : (
                      <span className="inline-flex items-center gap-1 rounded-full bg-amber-500/20 border border-amber-500/40 px-2 py-0.5 text-[11px] font-semibold text-amber-300">
                        <AlertTriangle className="h-3 w-3" />
                        Lệch Cặp ({activePair.diffLR > 0 ? `+${activePair.diffLR} Trái` : `${activePair.diffLR} Phải`})
                      </span>
                    )}
                  </div>
                  <h3 className="mt-1 text-sm font-bold text-white line-clamp-1">
                    {activePair.description}
                  </h3>
                </div>
                <div className="text-right">
                  <div className="text-[11px] text-slate-400">Bộ Ghép Đôi Hoàn Chỉnh</div>
                  <div className="text-lg font-extrabold font-mono text-cyan-300">
                    {activePair.matchedSets.toLocaleString('en-US')}{' '}
                    <span className="text-xs font-normal text-slate-400">bộ</span>
                  </div>
                </div>
              </div>

              {/* Left vs Right Detailed Comparison Cards */}
              <div className="mt-3 grid grid-cols-2 gap-2.5">
                {/* Left Side Box */}
                <div className="rounded-xl border border-cyan-500/30 bg-cyan-950/20 p-2.5">
                  <div className="flex items-center justify-between text-xs">
                    <span className="font-bold text-cyan-300">BÊN TRÁI (LHF)</span>
                    <span className="font-mono text-[11px] text-slate-300">{activePair.leftCode}</span>
                  </div>
                  <div className="mt-1.5 text-xl font-extrabold font-mono text-white">
                    {activePair.leftTotal.toLocaleString('en-US')}{' '}
                    <span className="text-xs font-normal text-slate-400">pcs</span>
                  </div>
                  <div className="mt-1.5 flex items-center justify-between text-[11px] text-slate-300 border-t border-cyan-500/20 pt-1.5 font-mono">
                    <span>Kho 01: {activePair.leftWh01.toLocaleString('en-US')}</span>
                    <span>Kho 60: {activePair.leftWh60.toLocaleString('en-US')}</span>
                  </div>
                  {activeLeftDemand && (
                    <div
                      className={`mt-1.5 rounded-md px-2 py-1 text-[11px] font-mono ${
                        activeLeftDemand.shortageQty > 0
                          ? 'bg-rose-500/20 text-rose-300 border border-rose-500/30'
                          : 'bg-emerald-500/15 text-emerald-300'
                      }`}
                    >
                      Cầu xuất: {activeLeftDemand.totalRequiredQty.toLocaleString('en-US')} |{' '}
                      {activeLeftDemand.shortageQty > 0
                        ? `Thiếu: -${activeLeftDemand.shortageQty.toLocaleString('en-US')}`
                        : `Dư: +${activeLeftDemand.surplusQty.toLocaleString('en-US')}`}
                    </div>
                  )}
                </div>

                {/* Right Side Box */}
                <div className="rounded-xl border border-emerald-500/30 bg-emerald-950/20 p-2.5">
                  <div className="flex items-center justify-between text-xs">
                    <span className="font-bold text-emerald-300">BÊN PHẢI (RHF)</span>
                    <span className="font-mono text-[11px] text-slate-300">{activePair.rightCode}</span>
                  </div>
                  <div className="mt-1.5 text-xl font-extrabold font-mono text-white">
                    {activePair.rightTotal.toLocaleString('en-US')}{' '}
                    <span className="text-xs font-normal text-slate-400">pcs</span>
                  </div>
                  <div className="mt-1.5 flex items-center justify-between text-[11px] text-slate-300 border-t border-emerald-500/20 pt-1.5 font-mono">
                    <span>Kho 01: {activePair.rightWh01.toLocaleString('en-US')}</span>
                    <span>Kho 60: {activePair.rightWh60.toLocaleString('en-US')}</span>
                  </div>
                  {activeRightDemand && (
                    <div
                      className={`mt-1.5 rounded-md px-2 py-1 text-[11px] font-mono ${
                        activeRightDemand.shortageQty > 0
                          ? 'bg-rose-500/20 text-rose-300 border border-rose-500/30'
                          : 'bg-emerald-500/15 text-emerald-300'
                      }`}
                    >
                      Cầu xuất: {activeRightDemand.totalRequiredQty.toLocaleString('en-US')} |{' '}
                      {activeRightDemand.shortageQty > 0
                        ? `Thiếu: -${activeRightDemand.shortageQty.toLocaleString('en-US')}`
                        : `Dư: +${activeRightDemand.surplusQty.toLocaleString('en-US')}`}
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* Scrollable Synchronized List of All Displayed L/R Pairs */}
          <div className="px-4 py-2 bg-slate-900/90 border-b border-slate-800/80 flex items-center justify-between text-xs font-semibold text-slate-300">
            <span className="inline-flex items-center gap-1.5">
              <Scale className="h-3.5 w-3.5 text-cyan-400" />
              Danh Sách Cặp Mã Trái / Phải ({displayedPairs.length})
            </span>
            <span className="text-[11px] text-slate-400 font-mono">Trái (LHF) vs Phải (RHF)</span>
          </div>

          <div className="flex-1 overflow-y-auto divide-y divide-slate-800/60">
            {displayedPairs.map((pair) => {
              const isSel = selectedModel === pair.model;
              const maxSide = Math.max(1, pair.leftTotal, pair.rightTotal);
              const lPct = Math.round((pair.leftTotal / maxSide) * 100);
              const rPct = Math.round((pair.rightTotal / maxSide) * 100);
              return (
                <div
                  key={pair.pairKey}
                  onClick={() => onSelectModel(isSel ? null : pair.model)}
                  onMouseEnter={() => setHoveredPair(pair)}
                  onMouseLeave={() => setHoveredPair(null)}
                  className={`px-4 py-2.5 transition cursor-pointer ${
                    isSel
                      ? 'bg-cyan-500/15 border-l-4 border-l-cyan-400'
                      : 'hover:bg-slate-800/60'
                  }`}
                >
                  <div className="flex items-center justify-between text-xs">
                    <div className="flex items-center gap-2">
                      <span className="font-mono font-bold text-white">#{pair.model}</span>
                      <span className="text-slate-400 truncate max-w-[140px]">
                        {pair.description}
                      </span>
                    </div>
                    <div className="flex items-center gap-1.5 font-mono text-xs">
                      <span className="text-cyan-300 font-semibold">
                        {pair.leftTotal.toLocaleString('en-US')}L
                      </span>
                      <span className="text-slate-600">/</span>
                      <span className="text-emerald-300 font-semibold">
                        {pair.rightTotal.toLocaleString('en-US')}R
                      </span>
                      {!pair.isBalanced && (
                        <span className="ml-1 rounded bg-amber-500/20 px-1.5 py-0.5 text-[10px] font-bold text-amber-300">
                          {pair.diffLR > 0 ? `+${pair.diffLR}L` : `${pair.diffLR}R`}
                        </span>
                      )}
                      {pair.hasWarehouseCrossBalance && (
                        <span className="ml-1 rounded bg-cyan-500/20 px-1.5 py-0.5 text-[10px] font-bold text-cyan-300">
                          Gộp 01+60
                        </span>
                      )}
                    </div>
                  </div>

                  {/* Dual Butterfly Progress Bar */}
                  <div className="mt-1.5 grid grid-cols-2 gap-1.5 items-center">
                    <div className="h-1.5 w-full rounded-full bg-slate-800 overflow-hidden flex justify-end">
                      <div
                        className="h-full rounded-full bg-gradient-to-l from-cyan-400 to-blue-500"
                        style={{ width: `${lPct}%` }}
                      />
                    </div>
                    <div className="h-1.5 w-full rounded-full bg-slate-800 overflow-hidden">
                      <div
                        className="h-full rounded-full bg-gradient-to-r from-emerald-400 to-teal-500"
                        style={{ width: `${rPct}%` }}
                      />
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </section>
  );
};
