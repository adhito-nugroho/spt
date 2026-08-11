                </div>
            </main>

            <!-- Footer minimal -->
            <footer class="px-8 py-2.5 border-t border-gray-100 bg-white flex-shrink-0">
                <p class="text-xs text-gray-300 text-right">SIPENSURAT v1.0</p>
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