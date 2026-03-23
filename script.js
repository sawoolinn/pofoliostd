document.addEventListener('DOMContentLoaded', () => {
    // 1. Navbar Scroll Effect
    const navbar = document.querySelector('.navbar');
    window.addEventListener('scroll', () => {
        if (window.scrollY > 50) {
            navbar.classList.add('scrolled');
        } else {
            navbar.classList.remove('scrolled');
        }
    });

    // 2. Mobile Menu Toggle
    const menuBtn = document.querySelector('.menu-btn');
    const navLinks = document.querySelector('.nav-links');
    
    menuBtn.addEventListener('click', () => {
        navLinks.classList.toggle('active');
        menuBtn.textContent = navLinks.classList.contains('active') ? '✕' : '☰';
    });

    // Close menu when clicking a link
    document.querySelectorAll('.nav-links a').forEach(link => {
        link.addEventListener('click', () => {
            navLinks.classList.remove('active');
            menuBtn.textContent = '☰';
        });
    });

    // 3. Scroll Reveal Animations using Intersection Observer
    const observerOptions = {
        root: null,
        rootMargin: '0px',
        threshold: 0.15
    };

    const observer = new IntersectionObserver((entries, observer) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('show-section');
                observer.unobserve(entry.target); // Stop observing once revealed
            }
        });
    }, observerOptions);

    document.querySelectorAll('.hidden-section').forEach(section => {
        observer.observe(section);
    });

    // 4. Contact Form Handling with Error Prevention and Fetch API
    const contactForm = document.getElementById('contact-form');
    const formStatus = document.getElementById('form-status');
    const submitBtn = document.getElementById('submit-btn');

    if (contactForm) {
        contactForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            
            // Basic validation
            const name = document.getElementById('name').value.trim();
            const email = document.getElementById('email').value.trim();
            const message = document.getElementById('message').value.trim();

            if (!name || !email || !message) {
                showFormStatus('Please fill in all fields.', 'error');
                return;
            }

            // UI feedback during submission
            const originalBtnText = submitBtn.textContent;
            submitBtn.textContent = 'Sending...';
            submitBtn.disabled = true;
            formStatus.textContent = '';

            try {
                // Send data to backend Express API
                const response = await fetch('/api/contact', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ name, email, message })
                });

                const data = await response.json();

                if (response.ok) {
                    showFormStatus(data.message || 'Message sent successfully!', 'success');
                    contactForm.reset();
                } else {
                    showFormStatus(data.message || 'Error sending message. Please try again.', 'error');
                }
            } catch (error) {
                console.error('Error submitting form:', error);
                showFormStatus('A network error occurred. Please try again later.', 'error');
            } finally {
                // Restore button state
                submitBtn.textContent = originalBtnText;
                submitBtn.disabled = false;
            }
        });
    }

    function showFormStatus(msg, type) {
        formStatus.textContent = msg;
        formStatus.className = 'form-status'; // Reset classes
        formStatus.classList.add(type === 'success' ? 'status-success' : 'status-error');
        
        // Remove the message after 5 seconds
        setTimeout(() => {
            formStatus.textContent = '';
            formStatus.className = 'form-status';
        }, 5000);
    }

    // 5. Fetch Full Content from Excel CMS API (Now Static via GitHub Pages)
    async function loadContent() {
        try {
            const res = await fetch('content.xlsx');
            if (!res.ok) return;
            const arrayBuffer = await res.arrayBuffer();
            
            // Parse Excel file in the browser using SheetJS
            const wb = XLSX.read(arrayBuffer, { type: 'array' });
            const data = {};
            wb.SheetNames.forEach(sheetName => {
                data[sheetName.toLowerCase()] = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]);
            });
            
            // Render Stats
            if (data.stats && data.stats.length > 0) {
                const statIds = {
                    'GPA': 'gpa',
                    'Projects': 'projects',
                    'Volunteering': 'volunteering'
                };
                
                data.stats.forEach(row => {
                    const idPrefix = statIds[row.Metric];
                    if (idPrefix && document.getElementById(`stat-${idPrefix}-val`)) {
                        document.getElementById(`stat-${idPrefix}-val`).textContent = row.Value;
                        document.getElementById(`stat-${idPrefix}-desc`).textContent = row.Description;
                    }
                });
            }

            // Render About Paragraphs
            if (data.about && data.about.length > 0) {
                data.about.forEach(row => {
                    const pElem = document.getElementById(`about-${row.Key}`);
                    if (pElem) pElem.textContent = row.Text;
                });
            }

            // Render Timelines
            if (data.timeline && data.timeline.length > 0) {
                const academicsTimeline = document.getElementById('academics-timeline');
                const experienceTimeline = document.getElementById('experience-timeline');
                if (academicsTimeline) academicsTimeline.innerHTML = '';
                if (experienceTimeline) experienceTimeline.innerHTML = '';

                data.timeline.forEach(item => {
                    let displayTitle = item.Title;
                    let displayDesc = item.Description;
                    
                    // Specific logic for Academics: Highlight Major, put University as description
                    if (item.Section === 'Academics') {
                        displayTitle = item.Description; // The Major
                        displayDesc = item.Title; // The University
                    }

                    const html = `
                    <div class="timeline-item">
                        <div class="timeline-dot"></div>
                        <div class="timeline-content glass-card">
                            <h3>${displayTitle}</h3>
                            <span class="timeline-date">${item.Date}</span>
                            <p>${displayDesc}</p>
                        </div>
                    </div>`;
                    
                    if (item.Section === 'Academics' && academicsTimeline) {
                        academicsTimeline.innerHTML += html;
                    } else if (item.Section === 'Experience' && experienceTimeline) {
                        experienceTimeline.innerHTML += html;
                    }
                });
            }
        } catch(err) {
            console.error('Error loading content from Excel:', err);
        }
    }
    loadContent();
});
