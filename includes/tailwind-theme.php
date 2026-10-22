<?php
/**
 * Palet identitas SIPENSURAT CDK Bojonegoro.
 * Warna aksen diturunkan dari hijau tua pada logo Provinsi Jawa Timur.
 */
$sipensuratBrand = [
    '50' => '#f0fdf4',
    '100' => '#dcfce7',
    '200' => '#bbf7d0',
    '300' => '#86efac',
    '400' => '#4ade80',
    '500' => '#15803d',
    '600' => '#15803d',
    '700' => '#166534',
    '800' => '#14532d',
    '900' => '#052e16',
    '950' => '#022c11',
];

$sipensuratPrimary = $sipensuratBrand['600'];
$sipensuratPrimaryDark = $sipensuratBrand['700'];
$sipensuratPrimarySoft = $sipensuratBrand['100'];
$sipensuratPrimaryLight = $sipensuratBrand['50'];
$sipensuratPrimaryBorder = $sipensuratBrand['200'];
$sipensuratPrimaryRgb = '21, 128, 61';
?>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
    :root {
        --color-primary: <?= $sipensuratPrimary ?>;
        --color-primary-dark: <?= $sipensuratPrimaryDark ?>;
        --color-primary-soft: <?= $sipensuratPrimarySoft ?>;
        --color-primary-light: <?= $sipensuratPrimaryLight ?>;
        --color-primary-border: <?= $sipensuratPrimaryBorder ?>;
        --color-primary-rgb: <?= $sipensuratPrimaryRgb ?>;
        --color-primary-ring: rgba(<?= $sipensuratPrimaryRgb ?>, 0.16);
    }
</style>
<script>
    (function () {
        const brand = {
            50: <?= json_encode($sipensuratBrand['50']) ?>,
            100: <?= json_encode($sipensuratBrand['100']) ?>,
            200: <?= json_encode($sipensuratBrand['200']) ?>,
            300: <?= json_encode($sipensuratBrand['300']) ?>,
            400: <?= json_encode($sipensuratBrand['400']) ?>,
            500: <?= json_encode($sipensuratBrand['500']) ?>,
            600: <?= json_encode($sipensuratBrand['600']) ?>,
            700: <?= json_encode($sipensuratBrand['700']) ?>,
            800: <?= json_encode($sipensuratBrand['800']) ?>,
            900: <?= json_encode($sipensuratBrand['900']) ?>,
            950: <?= json_encode($sipensuratBrand['950']) ?>
        };

        window.SIPENSURAT_THEME = {
            primary: <?= json_encode($sipensuratPrimary) ?>,
            primaryDark: <?= json_encode($sipensuratPrimaryDark) ?>,
            primarySoft: <?= json_encode($sipensuratPrimarySoft) ?>,
            primaryLight: <?= json_encode($sipensuratPrimaryLight) ?>,
            primaryRgb: <?= json_encode($sipensuratPrimaryRgb) ?>
        };

        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Inter', 'ui-sans-serif', 'system-ui', 'sans-serif'],
                    },
                    colors: {
                        primary: <?= json_encode($sipensuratPrimary) ?>,
                        'primary-dark': <?= json_encode($sipensuratPrimaryDark) ?>,
                        'primary-soft': <?= json_encode($sipensuratPrimarySoft) ?>,
                        'primary-light': <?= json_encode($sipensuratPrimaryLight) ?>,
                        'primary-border': <?= json_encode($sipensuratPrimaryBorder) ?>,
                        brand: brand,
                        // Alias kompatibilitas: utility indigo lama otomatis mengikuti palet hijau.
                        indigo: brand
                    }
                }
            }
        };
    })();
</script>
