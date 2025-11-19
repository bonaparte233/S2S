/**
 * S2S - Main JavaScript
 */

// Utility function to get CSRF token
function getCookie(name) {
    let cookieValue = null;
    if (document.cookie && document.cookie !== '') {
        const cookies = document.cookie.split(';');
        for (let i = 0; i < cookies.length; i++) {
            const cookie = cookies[i].trim();
            if (cookie.substring(0, name.length + 1) === (name + '=')) {
                cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                break;
            }
        }
    }
    return cookieValue;
}

// File upload label update
document.addEventListener('DOMContentLoaded', function() {
    // Update file upload labels when files are selected
    const fileInputs = document.querySelectorAll('input[type="file"]');
    
    fileInputs.forEach(input => {
        input.addEventListener('change', function() {
            const label = this.nextElementSibling;
            if (label && label.classList.contains('file-upload-label')) {
                if (this.files && this.files.length > 0) {
                    label.textContent = this.files[0].name;
                    label.style.borderColor = 'var(--success-color)';
                    label.style.backgroundColor = 'rgba(46, 204, 113, 0.05)';
                }
            }
        });
    });
    
    // Form validation
    const uploadForm = document.getElementById('uploadForm');
    if (uploadForm) {
        uploadForm.addEventListener('submit', function(e) {
            const docxInput = document.querySelector('input[name="docx_file"]');
            if (!docxInput || !docxInput.files || docxInput.files.length === 0) {
                e.preventDefault();
                alert('请选择讲稿文件！');
                return false;
            }
            
            const templateChoice = document.querySelector('input[name="template_choice"]:checked');
            if (templateChoice && templateChoice.value === 'upload') {
                const templateInput = document.querySelector('input[name="template_file"]');
                if (!templateInput || !templateInput.files || templateInput.files.length === 0) {
                    e.preventDefault();
                    alert('请上传自定义模板文件！');
                    return false;
                }
            }
        });
    }
});

// Auto-dismiss messages after 5 seconds
document.addEventListener('DOMContentLoaded', function() {
    const messages = document.querySelectorAll('.message');
    messages.forEach(message => {
        setTimeout(() => {
            message.style.opacity = '0';
            message.style.transition = 'opacity 0.5s ease';
            setTimeout(() => {
                message.remove();
            }, 500);
        }, 5000);
    });
});

// Smooth scroll for anchor links
document.addEventListener('DOMContentLoaded', function() {
    const anchorLinks = document.querySelectorAll('a[href^="#"]');
    anchorLinks.forEach(link => {
        link.addEventListener('click', function(e) {
            const targetId = this.getAttribute('href');
            if (targetId !== '#') {
                e.preventDefault();
                const targetElement = document.querySelector(targetId);
                if (targetElement) {
                    targetElement.scrollIntoView({
                        behavior: 'smooth',
                        block: 'start'
                    });
                }
            }
        });
    });
});

