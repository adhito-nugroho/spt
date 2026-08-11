                </div>
            </main>

            <!-- Footer -->
            <footer class="bg-white border-t border-gray-200 pt-3 mt-4 pb-6">
                <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col md:flex-row justify-between items-center gap-4">
                    <p class="text-slate-500 text-sm text-center md:text-left">
                        &copy; <?php echo date('Y'); ?> Sistem Pengelolaan Surat Tugas.
                    </p>
                    <p class="text-slate-400 text-xs text-center md:text-right">
                        Versi 1.0 | Dikembangkan oleh Tim SIPENSURAT
                    </p>
                </div>
            </footer>
        </div>
    </div>

    <script>
        // Mobile menu toggle
        const menuBtn = document.querySelector('button.md\\:hidden');
        const sidebar = document.querySelector('aside');
        
        if(menuBtn && sidebar) {
            menuBtn.addEventListener('click', () => {
                sidebar.classList.toggle('hidden');
                sidebar.classList.toggle('fixed');
                sidebar.classList.toggle('inset-0');
                sidebar.classList.toggle('w-full');
                sidebar.classList.toggle('z-50');
            });
        }
    </script>
</body>
</html>