/**
 * Pratinjau interaktif surat tugas.
 *
 * Halaman ini dirender server-side dari DOCX (modules/generate-surat.php
 * ?format=preview). Skrip ini menambahkan penggaris, pegangan jarak antar
 * blok, dan blok TTD yang bisa digeser, lalu mengirim perubahan ke panel
 * inspector lewat postMessage.
 */
(function () {
    'use strict';

    var CONFIG = window.SPT_PREVIEW_CONFIG || {};
    var settings = copy(CONFIG.settings || {});
    var ranges = CONFIG.ranges || {};
    var units = CONFIG.units || {};
    var pageWidthMm = Number(CONFIG.pageWidthMm) || 210;
    var pageHeightMm = Number(CONFIG.pageHeightMm) || 297;
    var dominantFont = Number(CONFIG.dominantFont) || 12;

    var pagesEl = document.getElementById('sptPages');
    if (!pagesEl) {
        return;
    }

    var rulerH = document.getElementById('sptRulerH');
    var rulerV = document.getElementById('sptRulerV');
    var gapHandles = [];
    var ttdMoves = [];
    var marginMarkers = {};
    var selectedEl = null;
    var repaginateTimer = null;
    var redrawTimer = null;

    var LABELS = {
        kop_bottom_spacing: 'Jarak bawah kop',
        nomor_surat_spacing: 'Judul \u2192 nomor',
        judul_bottom_spacing: 'Nomor \u2192 DASAR',
        dasar_item_spacing: 'Kerapatan butir DASAR',
        dasar_bottom_spacing: 'DASAR \u2192 MEMERINTAHKAN',
        memerintahkan_spacing: 'Spasi MEMERINTAHKAN',
        pegawai_gap_spacing: 'Jarak antar pegawai',
        untuk_top_spacing: 'Sebelum UNTUK',
        untuk_item_spacing: 'Kerapatan butir tugas',
        ttd_date_to_jabatan: 'Tanggal \u2192 jabatan',
        ttd_space_height: 'Ruang TTD',
        ttd_name_nip_spacing: 'Nama \u2192 NIP'
    };

    function copy(source) {
        var target = {};
        Object.keys(source || {}).forEach(function (key) {
            target[key] = source[key];
        });
        return target;
    }

    function pxPerMm() {
        return 96 / 25.4;
    }

    function zoom() {
        var value = parseFloat(pagesEl.style.zoom);
        return isFinite(value) && value > 0 ? value : 1;
    }

    function cssVarName(key) {
        return '--spt-' + key.replace(/_/g, '-');
    }

    function cssValue(key, value) {
        var unit = units[key];
        if (unit === undefined) {
            if (/^margin_/.test(key)) {
                unit = 'mm';
            } else if (key === 'font_size') {
                unit = 'pt';
            } else if (key === 'line_height') {
                unit = '';
            } else {
                unit = 'px';
            }
        }
        var number = Number(value);
        if (!isFinite(number)) {
            number = Number((CONFIG.settings || {})[key] || 0);
        }
        if (unit === '') {
            return String(number);
        }
        return trimNumber(number) + unit;
    }

    function trimNumber(number) {
        return String(Math.round(number * 100) / 100);
    }

    function clampValue(key, value) {
        var range = ranges[key];
        var number = Number(value);
        if (!isFinite(number)) {
            number = Number(settings[key] || 0);
        }
        if (!range) {
            return number;
        }
        number = Math.min(Number(range.max), Math.max(Number(range.min), number));
        var step = Number(range.step) || 1;
        var rounded = Math.round(number / step) * step;
        var decimals = (String(step).split('.')[1] || '').length;
        return parseFloat(rounded.toFixed(Math.max(decimals, 0)));
    }

    function displayValue(key) {
        var value = Number(settings[key]);
        var unit = units[key] === undefined ? 'px' : units[key];
        if (unit === '') {
            return trimNumber(value) + '\u00d7';
        }
        return trimNumber(value) + ' ' + unit;
    }

    function post(message) {
        if (window.parent === window) {
            return;
        }
        try {
            window.parent.postMessage(message, '*');
        } catch (error) {
            /* diabaikan */
        }
    }

    // ------------------------------------------------------------------
    // Terapkan pengaturan sebagai CSS variable.
    // ------------------------------------------------------------------
    function applySettings() {
        var root = document.documentElement;
        Object.keys(units).forEach(function (key) {
            if (settings[key] === undefined) {
                return;
            }
            root.style.setProperty(cssVarName(key), cssValue(key, settings[key]));
        });
        var factor = dominantFont > 0 ? Number(settings.font_size) / dominantFont : 1;
        if (!isFinite(factor) || factor <= 0) {
            factor = 1;
        }
        root.style.setProperty('--spt-font-factor', String(Math.round(factor * 10000) / 10000));
    }

    // ------------------------------------------------------------------
    // Pagination: pindahkan blok yang meluber ke halaman berikutnya.
    // ------------------------------------------------------------------
    function makePage() {
        var page = document.createElement('div');
        page.className = 'spt-page';
        var inner = document.createElement('div');
        inner.className = 'spt-page-inner';
        page.appendChild(inner);
        return page;
    }

    function availableHeight(page) {
        var style = window.getComputedStyle(page);
        return page.clientHeight - parseFloat(style.paddingTop || '0') - parseFloat(style.paddingBottom || '0');
    }

    function isOverlay(el) {
        var cls = el.classList;
        return !!cls && (cls.contains('spt-gap-handle') || cls.contains('spt-ttd-move') || cls.contains('spt-margin-marker'));
    }

    // Tinggi alur (in-flow) halaman. Kotak TTD berposisi absolut sehingga
    // scrollHeight ikut menghitungnya; itu membuat blok jangkar terlempar ke
    // halaman berikutnya padahal masih ada ruang.
    function flowHeight(inner) {
        var bottom = 0;
        Array.prototype.slice.call(inner.children).forEach(function (child) {
            if (isOverlay(child)) {
                return;
            }
            var style = window.getComputedStyle(child);
            if (style.position === 'absolute' || style.position === 'fixed') {
                return;
            }
            var margin = parseFloat(style.marginBottom) || 0;
            bottom = Math.max(bottom, child.offsetTop + child.offsetHeight + margin);
        });
        return bottom;
    }

    function repaginate() {
        normalizeTabs();
        var oldPages = Array.prototype.slice.call(pagesEl.querySelectorAll('.spt-page'));
        if (!oldPages.length) {
            return;
        }

        var blocks = [];
        oldPages.forEach(function (page, pageIndex) {
            var inner = page.querySelector('.spt-page-inner');
            if (!inner) {
                return;
            }
            var pageBlocks = Array.prototype.slice.call(inner.children).filter(function (child) {
                return !isOverlay(child);
            });
            pageBlocks.forEach(function (child, childIndex) {
                blocks.push({ el: child, force: pageIndex > 0 && childIndex === 0 });
            });
        });

        pagesEl.innerHTML = '';
        var page = makePage();
        pagesEl.appendChild(page);
        var inner = page.querySelector('.spt-page-inner');

        blocks.forEach(function (item) {
            if (item.force && inner.children.length > 0) {
                page = makePage();
                pagesEl.appendChild(page);
                inner = page.querySelector('.spt-page-inner');
            }
            inner.appendChild(item.el);

            if (inner.children.length > 1 && flowHeight(inner) > availableHeight(page) + 0.5) {
                var last = inner.lastElementChild;
                page = makePage();
                pagesEl.appendChild(page);
                inner = page.querySelector('.spt-page-inner');
                inner.appendChild(last);
            }
        });

        scheduleOverlayRebuild();
        postPageEstimate();
    }

    function scheduleRepaginate(delay) {
        if (repaginateTimer) {
            clearTimeout(repaginateTimer);
        }
        repaginateTimer = setTimeout(function () {
            repaginateTimer = null;
            repaginate();
        }, typeof delay === 'number' ? delay : 160);
    }

    // Word menggeser tab sesudah nomor butir ke indent kiri paragraf,
    // bukan selebar 0,5 inci seperti span tab statis.
    function normalizeTabs() {
        Array.prototype.slice.call(document.querySelectorAll('.spt-tab')).forEach(function (tab) {
            var paragraph = tab.parentElement;
            if (!paragraph) {
                return;
            }
            tab.style.width = '';

            var label = '';
            var prev = tab.previousSibling;
            var guard = 0;
            // Nomor butir bisa terpecah menjadi beberapa span ("3" + "."),
            // jadi label dirangkai mundur sampai karakter non-nomor muncul.
            while (prev && guard++ < 6 && label.length < 8) {
                label = (prev.textContent || '').replace(/\s+/g, '') + label;
                prev = prev.previousSibling;
                if (/[^\d.)]/.test(label)) {
                    break;
                }
            }
            if (!/^\d{1,2}[.)]?$/.test(label)) {
                return;
            }

            var style = window.getComputedStyle(paragraph);
            var hanging = parseFloat(style.textIndent) || 0;
            var left = parseFloat(style.marginLeft) || 0;
            if (hanging >= 0 || left <= 0) {
                return;
            }
            var current = tab.offsetLeft - paragraph.offsetLeft;
            // offsetLeft paragraf sudah memuat margin-left, jadi indent
            // lanjutan (target tab) berada di posisi 0 relatif paragraf.
            var width = Math.min(0.5 * 96, -current);
            if (width > 0) {
                tab.style.width = Math.round(width * 100) / 100 + 'px';
            }
        });
    }

    function postPageEstimate() {
        var count = pagesEl.querySelectorAll('.spt-page').length;
        post({ type: 'page-estimate', count: count > 0 ? count : 1 });
    }

    // ------------------------------------------------------------------
    // Overlay: pegangan jarak dan tombol geser TTD.
    // ------------------------------------------------------------------
    function clearOverlays() {
        gapHandles.forEach(function (handle) {
            if (handle.parentNode) {
                handle.parentNode.removeChild(handle);
            }
        });
        gapHandles = [];
        ttdMoves.forEach(function (handle) {
            if (handle.parentNode) {
                handle.parentNode.removeChild(handle);
            }
        });
        ttdMoves = [];
    }

    function offsetTop(target, container) {
        return target.getBoundingClientRect().top - container.getBoundingClientRect().top;
    }

    function updateHandleLabel(handle) {
        var key = handle.dataset.key;
        handle.pill.textContent = (LABELS[key] || key) + ' \u00b7 ' + displayValue(key);
    }

    function positionGap(handle) {
        if (!handle.target || !handle.container) {
            return;
        }
        handle.style.top = offsetTop(handle.target, handle.container) + 'px';
        updateHandleLabel(handle);
    }

    function addGap(target, key, container) {
        if (!target || !ranges[key]) {
            return;
        }
        var handle = document.createElement('div');
        handle.className = 'spt-gap-handle';
        handle.dataset.key = key;
        var pill = document.createElement('span');
        pill.className = 'spt-gap-pill';
        handle.appendChild(pill);
        handle.target = target;
        handle.container = container;
        handle.pill = pill;
        container.appendChild(handle);
        gapHandles.push(handle);
        positionGap(handle);
        bindGapDrag(handle);
    }

    function bindGapDrag(handle) {
        var key = handle.dataset.key;
        var startY = 0;
        var startValue = 0;
        var dragging = false;

        handle.addEventListener('pointerdown', function (event) {
            event.preventDefault();
            event.stopPropagation();
            dragging = true;
            startY = event.clientY;
            startValue = Number(settings[key]);
            handle.classList.add('is-dragging');
            try {
                handle.setPointerCapture(event.pointerId);
            } catch (error) {
                /* diabaikan */
            }
            document.body.style.cursor = 'row-resize';
            document.body.style.userSelect = 'none';
        });

        handle.addEventListener('pointermove', function (event) {
            if (!dragging) {
                return;
            }
            var delta = (event.clientY - startY) / zoom();
            settings[key] = clampValue(key, startValue + delta);
            applySettings();
            updateHandleLabel(handle);
            updateGapPositions();
            postSettings();
        });

        function stop() {
            if (!dragging) {
                return;
            }
            dragging = false;
            handle.classList.remove('is-dragging');
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            scheduleRepaginate();
        }

        handle.addEventListener('pointerup', stop);
        handle.addEventListener('pointercancel', stop);
    }

    function updateGapPositions() {
        gapHandles.forEach(positionGap);
    }

    function positionTtdMove(handle) {
        var boxRect = handle.box.getBoundingClientRect();
        var innerRect = handle.inner.getBoundingClientRect();
        handle.style.left = (boxRect.right - innerRect.left + 4) + 'px';
        handle.style.top = (boxRect.top - innerRect.top) + 'px';
    }

    function bindTtdDrag(handle) {
        var startX = 0;
        var startY = 0;
        var startTtdX = 0;
        var startOffset = 0;
        var dragging = false;

        handle.addEventListener('pointerdown', function (event) {
            event.preventDefault();
            event.stopPropagation();
            dragging = true;
            startX = event.clientX;
            startY = event.clientY;
            startTtdX = Number(settings.ttd_x);
            startOffset = Number(settings.ttd_offset);
            handle.classList.add('is-dragging');
            try {
                handle.setPointerCapture(event.pointerId);
            } catch (error) {
                /* diabaikan */
            }
            document.body.style.cursor = 'move';
            document.body.style.userSelect = 'none';
        });

        handle.addEventListener('pointermove', function (event) {
            if (!dragging) {
                return;
            }
            var dx = (event.clientX - startX) / zoom();
            var dy = (event.clientY - startY) / zoom();
            settings.ttd_x = clampValue('ttd_x', startTtdX + dx);
            settings.ttd_offset = clampValue('ttd_offset', startOffset + dy);
            applySettings();
            positionTtdMove(handle);
            updateGapPositions();
            postSettings();
        });

        function stop() {
            if (!dragging) {
                return;
            }
            dragging = false;
            handle.classList.remove('is-dragging');
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            scheduleRepaginate();
        }

        handle.addEventListener('pointerup', stop);
        handle.addEventListener('pointercancel', stop);
    }

    function attachTtdControls() {
        Array.prototype.slice.call(document.querySelectorAll('.spt-txbx-main')).forEach(function (box) {
            var inner = box.closest('.spt-page-inner');
            if (!inner) {
                return;
            }

            var move = document.createElement('div');
            move.className = 'spt-ttd-move';
            move.textContent = '\u2725';
            move.title = 'Geser blok tanda tangan';
            inner.appendChild(move);
            move.box = box;
            move.inner = inner;
            ttdMoves.push(move);
            positionTtdMove(move);
            bindTtdDrag(move);

            addGap(box.querySelector('.spt-ttd-jabatan'), 'ttd_date_to_jabatan', box);
            addGap(box.querySelector('.spt-ttd-name'), 'ttd_space_height', box);
            addGap(box.querySelector('.spt-ttd-nip'), 'ttd_name_nip_spacing', box);
        });
    }

    function rebuildOverlays() {
        clearOverlays();

        Array.prototype.slice.call(pagesEl.querySelectorAll('.spt-page-inner')).forEach(function (inner) {
            var first = function (selector) {
                return inner.querySelector(selector);
            };

            addGap(first('.spt-title'), 'kop_bottom_spacing', inner);
            addGap(first('.spt-nomor'), 'nomor_surat_spacing', inner);
            addGap(first('.spt-dasar'), 'judul_bottom_spacing', inner);
            addGap(first('.spt-dasar-item'), 'dasar_item_spacing', inner);
            addGap(first('.spt-memerintahkan'), 'dasar_bottom_spacing', inner);
            addGap(first('.spt-kepada'), 'memerintahkan_spacing', inner);
            addGap(first('.spt-untuk'), 'untuk_top_spacing', inner);

            var rows = inner.querySelectorAll('.spt-pegawai-row');
            for (var i = 1; i < rows.length; i++) {
                addGap(rows[i], 'pegawai_gap_spacing', inner);
            }

            var items = inner.querySelectorAll('.spt-untuk-item');
            for (var j = 1; j < items.length; j++) {
                addGap(items[j], 'untuk_item_spacing', inner);
            }
        });

        attachTtdControls();
        updateGapPositions();
        drawRulers();
    }

    function scheduleOverlayRebuild() {
        if (redrawTimer) {
            clearTimeout(redrawTimer);
        }
        redrawTimer = setTimeout(function () {
            redrawTimer = null;
            rebuildOverlays();
        }, 40);
    }

    function postSettings() {
        post({ type: 'settings-change', settings: copy(settings) });
    }

    // ------------------------------------------------------------------
    // Penggaris dan penanda margin.
    // ------------------------------------------------------------------
    function firstPageRect() {
        var page = pagesEl.querySelector('.spt-page');
        return page ? page.getBoundingClientRect() : null;
    }

    function drawRulers() {
        var pageRect = firstPageRect();
        var scale = pageRect && pageWidthMm > 0 ? pageRect.width / pageWidthMm : pxPerMm();

        drawRulerH(pageRect, scale);
        drawRulerV(scale);
        updateMarginMarkers(pageRect, scale);
    }

    function setupCanvas(canvas, cssWidth, cssHeight) {
        if (!canvas) {
            return null;
        }
        var ratio = window.devicePixelRatio || 1;
        canvas.width = Math.max(1, Math.round(cssWidth * ratio));
        canvas.height = Math.max(1, Math.round(cssHeight * ratio));
        var context = canvas.getContext('2d');
        context.setTransform(ratio, 0, 0, ratio, 0, 0);
        context.clearRect(0, 0, cssWidth, cssHeight);
        context.fillStyle = '#f8fafc';
        context.fillRect(0, 0, cssWidth, cssHeight);
        return context;
    }

    function drawRulerH(pageRect, scale) {
        var context = setupCanvas(rulerH, window.innerWidth, 22);
        if (!context) {
            return;
        }
        if (!pageRect) {
            return;
        }

        context.strokeStyle = '#94a3b8';
        context.fillStyle = '#64748b';
        context.font = '8px Arial';
        context.textAlign = 'center';
        context.textBaseline = 'top';

        for (var mm = 0; mm <= pageWidthMm; mm += 5) {
            var x = pageRect.left + mm * scale;
            if (x < 0 || x > window.innerWidth) {
                continue;
            }
            var major = mm % 10 === 0;
            context.beginPath();
            context.moveTo(Math.round(x) + 0.5, major ? 12 : 16);
            context.lineTo(Math.round(x) + 0.5, 22);
            context.stroke();
            if (major) {
                context.fillText(String(mm), x, 3);
            }
        }

        var left = pageRect.left + Number(settings.margin_left) * scale;
        var right = pageRect.left + pageRect.width - Number(settings.margin_right) * scale;
        context.strokeStyle = '#059669';
        context.lineWidth = 2;
        context.beginPath();
        context.moveTo(left, 10);
        context.lineTo(left, 22);
        context.moveTo(right, 10);
        context.lineTo(right, 22);
        context.stroke();
    }

    function drawRulerV(scale) {
        var context = setupCanvas(rulerV, 22, window.innerHeight);
        if (!context) {
            return;
        }

        context.strokeStyle = '#94a3b8';
        context.fillStyle = '#64748b';
        context.font = '8px Arial';
        context.textAlign = 'left';
        context.textBaseline = 'middle';
        context.lineWidth = 1;

        Array.prototype.slice.call(pagesEl.querySelectorAll('.spt-page')).forEach(function (page) {
            var rect = page.getBoundingClientRect();
            if (rect.bottom < 22 || rect.top > window.innerHeight) {
                return;
            }
            for (var mm = 0; mm <= pageHeightMm; mm += 5) {
                var y = rect.top + mm * scale;
                if (y < 22 || y > window.innerHeight) {
                    continue;
                }
                var major = mm % 10 === 0;
                context.beginPath();
                context.moveTo(major ? 12 : 16, Math.round(y) + 0.5);
                context.lineTo(22, Math.round(y) + 0.5);
                context.stroke();
                if (major) {
                    context.fillText(String(mm), 2, y);
                }
            }

            var top = rect.top + Number(settings.margin_top) * scale;
            var bottom = rect.top + rect.height - Number(settings.margin_bottom) * scale;
            context.strokeStyle = '#059669';
            context.lineWidth = 2;
            context.beginPath();
            context.moveTo(10, top);
            context.lineTo(22, top);
            context.moveTo(10, bottom);
            context.lineTo(22, bottom);
            context.stroke();
            context.strokeStyle = '#94a3b8';
            context.lineWidth = 1;
        });
    }

    function createMarginMarkers() {
        [['left', 'spt-mm-h'], ['right', 'spt-mm-h'], ['top', 'spt-mm-v'], ['bottom', 'spt-mm-v']].forEach(function (pair) {
            var side = pair[0];
            var marker = document.createElement('div');
            marker.className = 'spt-margin-marker ' + pair[1];
            marker.dataset.margin = side;
            marker.title = 'Geser margin ' + side;
            document.body.appendChild(marker);
            marginMarkers[side] = marker;
            bindMarginDrag(marker, side);
        });
    }

    function updateMarginMarkers(pageRect, scale) {
        if (!pageRect) {
            Object.keys(marginMarkers).forEach(function (side) {
                marginMarkers[side].style.display = 'none';
            });
            return;
        }
        var visible = pageRect.bottom > 22 && pageRect.top < window.innerHeight + pageRect.height;
        Object.keys(marginMarkers).forEach(function (side) {
            marginMarkers[side].style.display = visible ? '' : 'none';
        });

        var left = pageRect.left + Number(settings.margin_left) * scale;
        var right = pageRect.left + pageRect.width - Number(settings.margin_right) * scale;
        var top = pageRect.top + Number(settings.margin_top) * scale;
        var bottom = pageRect.top + pageRect.height - Number(settings.margin_bottom) * scale;

        marginMarkers.left.style.left = Math.round(left - 2) + 'px';
        marginMarkers.left.style.top = '0px';
        marginMarkers.right.style.left = Math.round(right - 2) + 'px';
        marginMarkers.right.style.top = '0px';
        marginMarkers.top.style.left = '0px';
        marginMarkers.top.style.top = Math.round(top - 2) + 'px';
        marginMarkers.bottom.style.left = '0px';
        marginMarkers.bottom.style.top = Math.round(bottom - 2) + 'px';
    }

    function bindMarginDrag(marker, side) {
        var start = 0;
        var startValue = 0;
        var scale = pxPerMm();
        var dragging = false;

        marker.addEventListener('pointerdown', function (event) {
            event.preventDefault();
            event.stopPropagation();
            dragging = true;
            var pageRect = firstPageRect();
            scale = pageRect && pageWidthMm > 0 ? pageRect.width / pageWidthMm : pxPerMm();
            start = side === 'left' || side === 'right' ? event.clientX : event.clientY;
            startValue = Number(settings['margin_' + side]);
            try {
                marker.setPointerCapture(event.pointerId);
            } catch (error) {
                /* diabaikan */
            }
            document.body.style.userSelect = 'none';
        });

        marker.addEventListener('pointermove', function (event) {
            if (!dragging) {
                return;
            }
            var delta = side === 'left' || side === 'right'
                ? (event.clientX - start) / scale
                : (event.clientY - start) / scale;
            if (side === 'right' || side === 'bottom') {
                delta = -delta;
            }
            settings['margin_' + side] = clampValue('margin_' + side, startValue + delta);
            applySettings();
            drawRulers();
            postSettings();
        });

        function stop() {
            if (!dragging) {
                return;
            }
            dragging = false;
            document.body.style.userSelect = '';
            scheduleRepaginate();
        }

        marker.addEventListener('pointerup', stop);
        marker.addEventListener('pointercancel', stop);
    }

    // ------------------------------------------------------------------
    // Pemilihan blok.
    // ------------------------------------------------------------------
    function selectBlock(blockName) {
        if (selectedEl) {
            selectedEl.classList.remove('is-selected');
            selectedEl = null;
        }
        if (!blockName) {
            return;
        }
        var target = pagesEl.querySelector('[data-spt-block="' + blockName + '"]');
        if (!target) {
            return;
        }
        selectedEl = target;
        target.classList.add('is-selected');
        try {
            target.scrollIntoView({ block: 'center', behavior: 'smooth' });
        } catch (error) {
            target.scrollIntoView();
        }
    }

    document.addEventListener('click', function (event) {
        if (event.target.closest('.spt-gap-handle, .spt-ttd-move, .spt-margin-marker')) {
            return;
        }
        var block = event.target.closest('[data-spt-block]');
        if (!block) {
            if (selectedEl) {
                selectedEl.classList.remove('is-selected');
                selectedEl = null;
            }
            return;
        }
        if (selectedEl) {
            selectedEl.classList.remove('is-selected');
        }
        selectedEl = block;
        block.classList.add('is-selected');
        post({ type: 'block-select', block: block.getAttribute('data-spt-block') });
    });

    // ------------------------------------------------------------------
    // Pesan dari panel inspector.
    // ------------------------------------------------------------------
    window.addEventListener('message', function (event) {
        if (window.parent === window || event.source !== window.parent) {
            return;
        }
        var data = event.data || {};
        if (data.type === 'set-settings' && data.settings) {
            Object.keys(data.settings).forEach(function (key) {
                if (key !== 'lock_ttd' && ranges[key]) {
                    settings[key] = clampValue(key, data.settings[key]);
                }
            });
            applySettings();
            updateGapPositions();
            drawRulers();
            scheduleRepaginate(120);
        } else if (data.type === 'highlight-block') {
            selectBlock(data.block);
        }
    });

    // ------------------------------------------------------------------
    // Zoom: kecilkan halaman bila viewport iframe lebih sempit dari A4.
    // ------------------------------------------------------------------
    function applyZoom() {
        var baseWidth = pageWidthMm * pxPerMm();
        var available = window.innerWidth - 86;
        var value = 1;
        if (available < baseWidth) {
            value = Math.max(0.5, available / baseWidth);
        }
        pagesEl.style.zoom = String(Math.round(value * 100) / 100);
    }

    function onViewportChange() {
        applyZoom();
        drawRulers();
    }

    window.addEventListener('resize', function () {
        onViewportChange();
        updateGapPositions();
    });
    window.addEventListener('scroll', drawRulers, { passive: true });

    // ------------------------------------------------------------------
    // Inisialisasi.
    // ------------------------------------------------------------------
    applySettings();
    applyZoom();
    createMarginMarkers();
    repaginate();

    if (document.fonts && document.fonts.ready && typeof document.fonts.ready.then === 'function') {
        document.fonts.ready.then(function () {
            scheduleRepaginate(0);
        });
    }
    window.addEventListener('load', function () {
        scheduleRepaginate(0);
    });

    post({ type: 'ready', count: pagesEl.querySelectorAll('.spt-page').length });
})();
