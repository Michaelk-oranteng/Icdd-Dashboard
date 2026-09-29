// responsive-sidebar.js
(function() {
    'use strict';
    
    const MOBILE_BREAKPOINT = 768;
    const TABLET_BREAKPOINT = 1024;
    
    class ResponsiveSidebar {
        constructor() {
            this.sidebar = document.querySelector('.nav-bar');
            this.mainContent = document.querySelector('.main-content-wrapper, .main-content');
            this.toggleBtn = document.querySelector('.sidebar-toggle');
            this.overlay = null;
            
            if (!this.sidebar) return;
            
            this.init();
        }
        
        init() {
            this.createOverlay();
            this.bindEvents();
            this.handleResize();
        }
        
        createOverlay() {
            this.overlay = document.createElement('div');
            this.overlay.className = 'sidebar-overlay';
            this.overlay.style.cssText = `
                position: fixed;
                inset: 0;
                background: rgba(15, 27, 51, 0.45);
                backdrop-filter: blur(4px);
                z-index: 999;
                opacity: 0;
                visibility: hidden;
                transition: all 0.3s ease;
            `;
            document.body.appendChild(this.overlay);
            
            // Add styles for visible state
            const style = document.createElement('style');
            style.textContent = `
                .sidebar-overlay.visible {
                    opacity: 1 !important;
                    visibility: visible !important;
                }
            `;
            document.head.appendChild(style);
        }
        
        bindEvents() {
            // Toggle button
            if (this.toggleBtn) {
                this.toggleBtn.addEventListener('click', () => this.toggle());
            }
            
            // Overlay click
            this.overlay.addEventListener('click', () => this.close());
            
            // Escape key
            document.addEventListener('keydown', (e) => {
                if (e.key === 'Escape') this.close();
            });
            
            // Resize with debounce
            let resizeTimer;
            window.addEventListener('resize', () => {
                clearTimeout(resizeTimer);
                resizeTimer = setTimeout(() => this.handleResize(), 100);
            });
            
            // Touch swipe to close
            let touchStartX = 0;
            this.sidebar.addEventListener('touchstart', (e) => {
                touchStartX = e.touches[0].clientX;
            }, { passive: true });
            
            this.sidebar.addEventListener('touchend', (e) => {
                const touchEndX = e.changedTouches[0].clientX;
                const diff = touchStartX - touchEndX;
                
                // Swipe left to close
                if (diff > 50 && this.isMobile()) {
                    this.close();
                }
            }, { passive: true });
        }
        
        isMobile() {
            return window.innerWidth < MOBILE_BREAKPOINT;
        }
        
        isTablet() {
            return window.innerWidth >= MOBILE_BREAKPOINT && 
                   window.innerWidth < TABLET_BREAKPOINT;
        }
        
        handleResize() {
            if (this.isMobile()) {
                // Mobile: sidebar is hidden by default, shown as overlay
                this.sidebar.classList.remove('open');
                document.body.classList.remove('sidebar-collapsed');
                this.close();
            } else if (this.isTablet()) {
                // Tablet: sidebar is collapsed by default
                document.body.classList.add('sidebar-collapsed');
                this.sidebar.classList.remove('open');
                this.overlay.classList.remove('visible');
            } else {
                // Desktop: restore saved state or default to expanded
                const savedState = localStorage.getItem('sidebarCollapsed');
                if (savedState === 'true') {
                    document.body.classList.add('sidebar-collapsed');
                } else {
                    document.body.classList.remove('sidebar-collapsed');
                }
                this.sidebar.classList.remove('open');
                this.overlay.classList.remove('visible');
            }
        }
        
        toggle() {
            if (this.isMobile()) {
                if (this.sidebar.classList.contains('open')) {
                    this.close();
                } else {
                    this.open();
                }
            } else {
                // Desktop/Tablet: toggle collapsed state
                document.body.classList.toggle('sidebar-collapsed');
                const isCollapsed = document.body.classList.contains('sidebar-collapsed');
                localStorage.setItem('sidebarCollapsed', isCollapsed);
            }
        }
        
        open() {
            this.sidebar.classList.add('open');
            this.overlay.classList.add('visible');
            document.body.style.overflow = 'hidden';
        }
        
        close() {
            this.sidebar.classList.remove('open');
            this.overlay.classList.remove('visible');
            document.body.style.overflow = '';
        }
    }
    
    // Initialize when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => new ResponsiveSidebar());
    } else {
        new ResponsiveSidebar();
    }
})();