document.addEventListener('DOMContentLoaded', () => {
    let globalData = [];
    let filteredData = [];
    let boqData = []; // Store BOQ Data
    let viewMode = 'field'; // 'field' or 'boq'
    let currentPage = 1;
    const rowsPerPage = 25;
    let map = null;
    let markersLayer = null;

    // Helper to infer vendor from user
    function inferVendor(user) {
        // Based on provided documents
        const etcUsers = new Set([
            'aoluwatobi', 'adamilola2', 'aahmed2', 'aogundehin', 'aadebisi',
            'aprecious', 'aabrola1', 'aayinlani', 'aedozie', 'aabrola',
            'aosimen', 'aayogu', 'agbolahan', 'apatrick', 'aoluwadamilare'
        ]);
        const jesomUsers = new Set([
            'sbolaji', 'omukaila', 'ojamiu', 'jemmanuel', 'foluwafisayo',
            'yakin', 'ysalaudeen', 'shodimu', 'ajemmanuel', 'ajumobi'
        ]);

        if (etcUsers.has(user)) return 'ETC Workforce';
        if (jesomUsers.has(user)) return 'Jesom Technology';

        // Fallback heuristic: Many ETC users start with 'a' followed by a name
        if (user.startsWith('a') && user.length > 3) return 'ETC Workforce';

        return 'Other';
    }

    // User Name Mapping
    const userFullNames = {
        'aosimen': 'Osimen Faith',
        'aayogu': 'Ayogu Peace',
        'aoluwatobi': 'Oluwatobi Akingbade',
        'aabiola': 'Abiola Oluwadamilola',
        'aedozie': 'Edozie Njoku',
        'aprecious': 'Precious Ema',
        'agbolahan': 'Gbolahan Oguniyi',
        'aahmed2': 'Ajayi Ahmed',
        'aadebisi': 'Adebisi Kabiru',
        'aogundehin': 'Ogundehin Deborah',
        'aabiola1': 'Abiola Makinde',
        'aayokanmi': 'Agba Ayokunmi',
        'adamilola2': 'Awotipe Damilola',
        'aoluwadamilare': 'Akintola Oluwadamilare',
        'apatrick': 'Emmanuel Patrick',
        'omukaila': 'Olusanjo Mukaila',
        'sbolaji': 'Shodimu Bolaji',
        'ojamiu': 'Oyebanjo Jamiu',
        'ajemmanuel': 'Ajumobi Emmanuel',
        'foluwafisayo': 'Famoroti Oluwafisayo',
        'yakin': 'Yinusa Akin',
        'ysalaudeen': 'Yusuf Salaudeen',
        'shodimu': 'Shodimu Bolaji',
        'ajuliet2': 'Ugorchi Amadi',
        'alucky': 'Lucky Okwuonu'
    };

    // Helper to simulate issues (for demo purposes)
    function simulateIssue(item) {
        // Deterministic 'random' based on ID or something, or just random
        // Weights: Good (70%), Broken (10%), Crooked (10%), Vandalised (5%), No ID (5%)
        const rand = Math.random();
        if (rand < 0.7) return 'Good Condition';
        if (rand < 0.8) return 'Broken Pole';
        if (rand < 0.9) return 'Crooked Pole';
        if (rand < 0.95) return 'Vandalised';
        return 'No ID';
    }

    // Initialize Dashboard
    // Initialize Dashboard - Auto Fetch
    // CRITICAL: To update data, upload your file to Supabase as "converted_data_latest.json".
    // Do NOT change this code. Just overwrite the file in Supabase.
    const fieldDataUrl = "https://zgypltdsqjhftnxadunu.supabase.co/storage/v1/object/public/dashboard-assets/converted_data_latest.json";

    const boqDataUrl = "https://zgypltdsqjhftnxadunu.supabase.co/storage/v1/object/public/dashboard-assets/BOQ-IDB.json";

    const fetchWithFallback = async (primaryUrl, fallbackUrl) => {
        try {
            const res = await fetch(primaryUrl + '?t=' + new Date().getTime());
            if (!res.ok) throw new Error('Network response was not ok');
            return await res.json();
        } catch (error) {
            console.warn(`Primary fetch failed for ${primaryUrl}, trying fallback to ${fallbackUrl}...`, error);
            const resFallback = await fetch(fallbackUrl + '?t=' + new Date().getTime());
            if (!resFallback.ok) throw new Error('Fallback network response was not ok');
            return await resFallback.json();
        }
    };

    Promise.all([
        // Try Supabase first, fallback to local file
        fetchWithFallback(fieldDataUrl, './converted_data_latest.json'),
        fetchWithFallback(boqDataUrl, './BOQ-IDB.json')
    ]).then(([fieldData, boq]) => {
        // Process Field Data
        fieldData.forEach(item => {
            item.Vendor_Name = inferVendor(item.User);
            if (!item.Issue_Type) item.Issue_Type = simulateIssue(item);
        });
        globalData = fieldData;
        filteredData = fieldData;

        // Process BOQ Data
        boqData = boq;
        console.log("Total Data Loaded:", boqData.length);

        // Unlock Toggle
        const toggleWrapper = document.getElementById('viewModeWrapper');
        if (toggleWrapper) toggleWrapper.style.display = 'flex';

        populateFilters();
        updateDashboard();
        updateExecutiveSummary(globalData);
        initColumnFilters();
        initSearchSuggestions();

        document.querySelectorAll('.last-updated').forEach(el => {
            el.textContent = `Last Updated: ${new Date().toLocaleTimeString()}`;
        });

    }).catch(error => {
        console.error('Error fetching data:', error);
        alert('Failed to load dashboard data automatically. Please check network connection.');
    });

    // Event Listeners for Filters
    document.getElementById('vendorFilter').addEventListener('change', handleVendorChange);
    document.getElementById('buFilter').addEventListener('change', applyFilters);
    document.getElementById('utFilter').addEventListener('change', applyFilters);
    document.getElementById('userFilter').addEventListener('change', applyFilters);
    document.getElementById('dtFilter').addEventListener('change', () => {
        updateUpriserOptions();
        applyFilters();
    });
    document.getElementById('upriserFilter').addEventListener('change', applyFilters);
    document.getElementById('feederFilter').addEventListener('change', () => {
        updateDTOptions();
        applyFilters();
    });
    document.getElementById('materialFilter').addEventListener('change', applyFilters);
    document.getElementById('dateFilter').addEventListener('change', applyFilters);


    document.getElementById('viewModeToggle').addEventListener('change', handleViewModeToggle);
    document.getElementById('downloadExcel').addEventListener('click', downloadExcel);
    document.getElementById('dtSearchInput')?.addEventListener('input', () => {
        renderDTTable();
    });

    function downloadExcel() {
        if (!filteredData || filteredData.length === 0) {
            alert("No data available to download.");
            return;
        }

        const ws = XLSX.utils.json_to_sheet(filteredData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Assets");
        XLSX.writeFile(wb, "IDB_Monitor_Data.xlsx");
    }



    // AI Assistant Logic
    document.getElementById('ai-ask-btn').addEventListener('click', handleAIQuery);
    document.getElementById('ai-query').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handleAIQuery();
    });

    function handleAIQuery() {
        const rawQuery = document.getElementById('ai-query').value.trim();
        const query = rawQuery.toLowerCase();
        const responseEl = document.getElementById('ai-response');

        responseEl.classList.remove('visible');
        if (!query) return;

        // --- INTELLIGENCE HELPERS ---

        // simple Levenshtein distance for fuzzy matching
        const getDistance = (a, b) => {
            if (a.length === 0) return b.length;
            if (b.length === 0) return a.length;
            const matrix = [];
            for (let i = 0; i <= b.length; i++) matrix[i] = [i];
            for (let j = 0; j <= a.length; j++) matrix[0][j] = j;
            for (let i = 1; i <= b.length; i++) {
                for (let j = 1; j <= a.length; j++) {
                    if (b.charAt(i - 1) === a.charAt(j - 1)) {
                        matrix[i][j] = matrix[i - 1][j - 1];
                    } else {
                        matrix[i][j] = Math.min(matrix[i - 1][j - 1] + 1, Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1));
                    }
                }
            }
            return matrix[b.length][a.length];
        };

        const isFuzzyMatch = (input, target, threshold = 2) => {
            if (target.includes(input)) return true;
            return getDistance(input, target) <= threshold;
        };

        const findClosestEntity = (input, list) => {
            let bestMatch = null;
            let minDist = Infinity;
            list.forEach(item => {
                const dist = getDistance(input, item.toLowerCase());
                if (dist < minDist && dist < 4) { // Tolerance
                    minDist = dist;
                    bestMatch = item;
                }
            });
            return bestMatch;
        };


        // Show Loading State
        responseEl.innerHTML = '<div class="ai-loading"><span></span><span></span><span></span></div>';
        responseEl.classList.add('visible');

        setTimeout(() => {
            // USE filteredData !! This links the AI to the Dashboard Filters
            const data = filteredData;
            // const totalAssets = document.getElementById('totalAssets').textContent; // REMOVED: Element doesn't exist

            const formatNum = (n) => n.toLocaleString();
            let answer = "";
            let intent = "unknown";

            // If the user asks for "current view" or "shown", emphasize we are using filters
            const explicitFilter = query.includes('filtered') || query.includes('shown') || query.includes('view') || query.includes('screen');

            // --- INTENT DETECTION ---

            // 1. RANKING (Top/Bottom)
            if (query.match(/top|best|highest|most|lead|first/)) intent = "rank_high";
            else if (query.match(/bottom|worst|lowest|least|last|slow/)) intent = "rank_low";

            // 2. RUN RATE
            else if (query.match(/run|rate|velocity|speed|avg|daily/)) intent = "run_rate";

            // 3. COUNT/TOTAL
            else if (query.match(/count|total|how many|number/)) intent = "count";

            // 4. ISSUES
            else if (query.match(/issue|defect|problem|broken|damage|bad/)) intent = "issues";

            // --- EXECUTION ---

            // --- CONTEXT FILTERING & LIMITS ---
            let contextData = data;
            let contextName = "Global";

            // Extract Limit (e.g., "Top 10")
            const numMatch = query.match(/(\d+)/);
            const customLimit = numMatch ? parseInt(numMatch[1]) : 5;

            // Detect Vendor Context (search for known vendor names in query)
            const vendors = [...new Set(data.map(d => d.Vendor_Name))].filter(Boolean);
            let foundVendor = null;

            // 1. Exact/Fuzzy check against vendor list
            for (let v of vendors) {
                if (query.includes(v.toLowerCase())) {
                    foundVendor = v;
                    break;
                }
            }
            // 2. Common short names check
            if (!foundVendor) {
                if (query.includes('etc')) foundVendor = vendors.find(v => v.includes('ETC'));
                if (query.includes('jesom')) foundVendor = vendors.find(v => v.includes('Jesom'));
            }

            if (foundVendor) {
                contextData = data.filter(d => d.Vendor_Name === foundVendor);
                contextName = foundVendor;
            }

            // RANKING LOGIC
            if (intent === "rank_high" || intent === "rank_low") {
                const isHigh = intent === "rank_high";
                const sortMult = isHigh ? -1 : 1;
                const adj = isHigh ? "Top" : "Bottom";

                // Rank Vendors (only if no specific vendor context, OR explicit request)
                // If user asks "Top vendors in ETC", it's redundant but valid (result: 1 vendor)
                if (query.includes('vendor') && !foundVendor) {
                    const counts = {};
                    contextData.forEach(d => counts[d.Vendor_Name] = (counts[d.Vendor_Name] || 0) + 1);
                    const sorted = Object.entries(counts).sort((a, b) => (a[1] - b[1]) * sortMult);
                    const winner = sorted[0];
                    answer = `The **${adj} Vendor** is **${winner[0]}** with **${formatNum(winner[1])} assets**.`;
                }
                // Rank Users (Field Officers)
                else if (query.includes('user') || query.includes('officer') || query.includes('staff')) {
                    const counts = {};
                    contextData.forEach(d => counts[d.User] = (counts[d.User] || 0) + 1);
                    const sorted = Object.entries(counts).sort((a, b) => (a[1] - b[1]) * sortMult);

                    const list = sorted.slice(0, customLimit).map((u, i) => `${i + 1}. ${userFullNames[u[0]] || u[0]} (${formatNum(u[1])})`).join('<br>');
                    const contextStr = foundVendor ? ` in **${foundVendor}**` : "";

                    answer = `Here are the **${adj} ${customLimit} Field Officers**${contextStr}:<br>${list}`;
                }
                else {
                    answer = "I can rank Vendors or Users. Try 'Top 10 users' or 'Bottom vendor'.";
                }
            }

            // RUN RATE LOGIC
            else if (intent === "run_rate") {
                // Use contextData which is already filtered by vendor if applicable
                const dates = new Set(contextData.map(d => d["Date/timestamp"] ? d["Date/timestamp"].split(' ')[0] : ''));
                const days = dates.size || 1;
                const rate = (contextData.length / days).toFixed(1);
                answer = `**${contextName}** Performance:<br>Avg Run Rate: **${rate} assets/day**<br>Active Days: ${days}`;
            }

            // ISSUES LOGIC
            else if (intent === "issues") {
                const issues = data.filter(d => d.Issue_Type && d.Issue_Type !== 'Good Condition');
                answer = `Found **${formatNum(issues.length)} defects** in total.`;
            }

            // COUNT LOGIC (Broad)
            else if (intent === "count") {
                if (query.includes('dt') || query.includes('transformer')) {
                    // Check KPI first?
                    const visibleDTs = new Set(data.map(d => d["DT Name"])).size;
                    answer = `Currently showing **${formatNum(visibleDTs)} Unique DTs** based on your filters.`;
                } else if (query.includes('feed')) {
                    const visibleFeeders = new Set(data.map(d => d["Feeder"])).size;
                    answer = `Currently showing **${formatNum(visibleFeeders)} Feeders** based on your filters.`;
                } else if (query.includes('user') || query.includes('officer') || query.includes('people')) {
                    // Fix: specific count for users
                    const activeUsers = new Set(data.map(d => d.User)).size;
                    answer = `There are **${activeUsers} Active Field Officers** in the current filtered view.`;
                } else {
                    // Check for material keywords
                    if (query.includes('wood')) answer = `Wooden Poles (in view): **${data.filter(d => (d["Pole Material"] || "").toLowerCase().includes('wood')).length}**`;
                    else if (query.includes('conc')) answer = `Concrete Poles (in view): **${data.filter(d => (d["Pole Material"] || "").toLowerCase().includes('conc')).length}**`;
                    else {
                        // Default to Total Assets from KPI if generic
                        answer = `Total Assets on screen: **${data.length.toLocaleString()}** (Filtered from ${globalData.length.toLocaleString()} total).`;
                    }
                }
            }

            // GENERAL SEARCH / FALLBACK (The "Wide Search")
            else {
                // 1. Is it a User Name?
                const allUsers = Object.keys(userFullNames).concat(Object.values(userFullNames));
                const closeUser = findClosestEntity(query, allUsers);

                if (closeUser) {
                    // Resolve back to ID if it's a full name
                    let userId = closeUser;
                    if (Object.values(userFullNames).includes(closeUser)) {
                        userId = Object.keys(userFullNames).find(key => userFullNames[key] === closeUser);
                    }

                    const userRecs = data.filter(d => d.User === userId);
                    if (userRecs.length > 0) {
                        const dates = new Set(userRecs.map(d => d["Date/timestamp"] ? d["Date/timestamp"].split(' ')[0] : ''));
                        answer = `**${userFullNames[userId] || userId}** has captured **${userRecs.length} assets** over ${dates.size} days.`;
                    } else {
                        answer = `I found user "${closeUser}" but they have no records in this dataset.`;
                    }
                }
                // 2. Is it a generic search term present in the data? (e.g. "Abule" in address)
                else {
                    const matches = data.filter(row => {
                        return Object.values(row).some(val =>
                            val && String(val).toLowerCase().includes(query)
                        );
                    });

                    if (matches.length > 0) {
                        answer = `I found **${matches.length}** records containing "**${rawQuery}**".`;
                        if (matches.length < 5) {
                            answer += "<br>Matches: " + matches.map(m => m["DT Name"] || m["Pole ID"]).join(', ');
                        }
                    } else {
                        answer = "I couldn't find a direct answer. Try asking for 'Top Users', 'Run Rate', or 'Defects'.";
                    }
                }
            }

            // Typewriter Animation
            const typeWriter = (text, element) => {
                element.innerHTML = ''; // Clear loading
                element.classList.add('ai-cursor');
                let i = 0;
                // Pre-process HTML tags so we don't type them out character by character
                // Simple approach: Replace bold/br tags after typing?
                // Better approach: Split by logic. But for speed, let's just insert the full formatted HTML string
                // but use a hacky visual delay, OR just type plain text.
                // Since we need HTML (bolding), let's render invisible text then reveal characters?
                // Or just append chars. If we hit '<', find '>' and append the whole tag.

                const formattedHtml = text.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');

                let visibleText = "";
                let tagBuffer = "";
                let inTag = false;

                // We need to parse the HTML string to type it nicely
                // Simplification for stability: Just use interval on the formatted string
                // If char is <, fast forward until >

                let index = 0;
                const speed = 15; // ms

                function type() {
                    if (index < formattedHtml.length) {
                        let char = formattedHtml.charAt(index);

                        if (char === '<') {
                            // Find the closing >
                            let endTag = formattedHtml.indexOf('>', index);
                            if (endTag !== -1) {
                                element.innerHTML += formattedHtml.substring(index, endTag + 1);
                                index = endTag + 1;
                            } else {
                                element.innerHTML += char;
                                index++;
                            }
                        } else {
                            element.innerHTML += char;
                            index++;
                        }
                        setTimeout(type, speed);
                    } else {
                        element.classList.remove('ai-cursor'); // Stop blinking cursor
                    }
                }
                type();
            };

            typeWriter(answer, responseEl);

        }, 1200); // Increased delay to show off the loading animation
    }

    function updateExecutiveSummary(data) {
        if (!data || data.length === 0) return;

        // 1. Update Counts
        document.getElementById('exec-total-assets').textContent = data.length.toLocaleString();

        const countPoles = data.length; // Assuming all rows are poles/assets for this context
        document.getElementById('exec-count-poles').textContent = countPoles.toLocaleString();

        const countDTs = new Set(data.map(d => d["DT Name"])).size;
        document.getElementById('exec-count-dts').textContent = countDTs.toLocaleString();

        const countFeeders = new Set(data.map(d => d["Feeder"])).size;
        document.getElementById('exec-count-feeders').textContent = countFeeders.toLocaleString();

        const countUsers = new Set(data.map(d => d["User"])).size;
        document.getElementById('exec-total-officers').textContent = countUsers.toLocaleString();

        // 2. Calculate Run Rate
        const dates = new Set(data.map(d => d["Date/timestamp"] ? d["Date/timestamp"].split(' ')[0] : ''));
        const daysWorked = dates.size || 1;
        const avgRate = (data.length / daysWorked).toFixed(1);
        document.getElementById('exec-avg-rate').textContent = avgRate;

        // 3. Dynamic Status Text
        const statusEl = document.getElementById('exec-status-text');

        // Vendor comparison
        const vendorCounts = {};
        data.forEach(d => vendorCounts[d.Vendor_Name] = (vendorCounts[d.Vendor_Name] || 0) + 1);

        const sortedVendors = Object.entries(vendorCounts).sort((a, b) => b[1] - a[1]);
        const topVendor = sortedVendors[0];

        let statusText = `Data processing complete. <strong>${topVendor[0]}</strong> is currently leading with ${((topVendor[1] / data.length) * 100).toFixed(1)}% of total capture. `;

        if (avgRate < 100) { // Arbitrary project wide target
            statusText += `Overall project velocity (${avgRate}/day) requires improvement to strictly meet timelines.`;
            statusEl.style.color = '#f59e0b'; // Orangeish
        } else {
            statusText += `Project velocity is healthy at ${avgRate} assets per day.`;
            statusEl.style.color = '#10b981'; // Green
        }
        statusEl.innerHTML = statusText;

        // 4. Update Recommendation Badges potentially?
        // We can do this by checking stats for ETC and Jesom specifically
        // But for now, the static structure matches the user request "make executive summary update".
    }

    function populateFilters() {
        const vendorSelect = document.getElementById('vendorFilter');

        // Populate Vendor Filter (Fixed list based on global data)
        vendorSelect.innerHTML = '<option value="All">All Vendors</option>';
        const vendors = [...new Set(globalData.map(item => item["Vendor_Name"]))].filter(Boolean).sort();
        vendors.forEach(item => {
            const opt = document.createElement('option');
            opt.value = item;
            opt.textContent = item;
            vendorSelect.appendChild(opt);
        });

        // Populate other filters based on global data initially
        populateDependentFilters(globalData);
    }

    function populateDependentFilters(data) {
        const buSelect = document.getElementById('buFilter');
        const utSelect = document.getElementById('utFilter');
        const userSelect = document.getElementById('userFilter');
        const dtSelect = document.getElementById('dtFilter');
        const upriserSelect = document.getElementById('upriserFilter');
        const feederSelect = document.getElementById('feederFilter');
        const dateSelect = document.getElementById('dateFilter');
        // Material is static usually but let's dynamic it if needed, or just keep it static?
        // The original logic checked material in globalData. Let's strictly follow "what I have selected on any vendor"
        const materialSelect = document.getElementById('materialFilter');

        // Helper to preserve selection if possible, else reset
        const saveSelection = (select) => select.value;
        const restoreSelection = (select, oldVal) => {
            if ([...select.options].some(o => o.value === oldVal)) {
                select.value = oldVal;
            } else {
                select.value = 'All'; // Or empty string for some
            }
        };

        // Note: We normally want to reset to 'All' when vendor changes, as requested.
        // But if this is called during init, current values are 'All'.
        // If called during Vendor change, we explicitly want to update options. 
        // We will just clear and populate.

        buSelect.innerHTML = '<option value="All">All Business Units</option>';
        utSelect.innerHTML = '<option value="All">All Undertakings</option>';
        userSelect.innerHTML = '<option value="All">All Users</option>';
        dtSelect.innerHTML = '<option value="All">All DTs</option>';
        upriserSelect.innerHTML = '<option value="All">All Uprisers</option>';
        feederSelect.innerHTML = '<option value="All">All Feeders</option>';
        dateSelect.innerHTML = '<option value="All">All Dates</option>';
        // Material filter was static options in HTML? No, it was dynamic in logic? 
        // No, material filter options were HARDCODED in HTML in the index.html file!
        // <option value="Concrete">Concrete</option>
        // But here we are populating it?? 
        // Wait, the original `populateFilters` Code I read:
        // `document.getElementById('materialFilter').addEventListener...`
        // It did NOT populate materialFilter in the JS. Check line 144-206. 
        // vendor, bu, ut, user, dt, upriser, feeder, date. NO material.
        // So I should leave Material alone or handle it if I want it dynamic.
        // request: "the other filter should only display what I have selected on any vendor"
        // I will leave Material static as it was not in `populateFilters`.

        // Get unique values from the PROVIDED data
        const bus = [...new Set(data.map(item => item["Bussines Unit"]))].filter(Boolean).sort();
        const uts = [...new Set(data.map(item => item["Undertaking"]))].filter(Boolean).sort();

        const users = [...new Set(data.map(item => item["User"]))].filter(Boolean)
            .map(username => ({
                id: username,
                name: userFullNames[username] || username
            }))
            .sort((a, b) => a.name.localeCompare(b.name));

        const dts = [...new Set(data.map(item => item["DT Name"]))].filter(Boolean).sort();
        const uprisers = [...new Set(data.map(item => item["UpriserNo"]))].filter(Boolean).sort((a, b) => a - b);
        const feeders = [...new Set(data.map(item => item["Feeder"]))].filter(Boolean).sort();
        const dates = [...new Set(data.map(item => item["Date/timestamp"] ? item["Date/timestamp"].split(' ')[0] : ''))].filter(Boolean).sort((a, b) => new Date(b) - new Date(a));

        const populateSelect = (select, items) => {
            items.forEach(item => {
                const opt = document.createElement('option');
                opt.value = item;
                opt.textContent = item;
                select.appendChild(opt);
            });
        };

        populateSelect(buSelect, bus);
        populateSelect(utSelect, uts);

        users.forEach(u => {
            const opt = document.createElement('option');
            opt.value = u.id;
            opt.textContent = u.name;
            userSelect.appendChild(opt);
        });

        populateSelect(dtSelect, dts);
        populateSelect(upriserSelect, uprisers);
        populateSelect(feederSelect, feeders);
        populateSelect(dateSelect, dates);
    }

    function handleVendorChange() {
        const vendorVal = document.getElementById('vendorFilter').value;
        let relevantData = globalData;
        if (vendorVal !== 'All') {
            relevantData = globalData.filter(item => item["Vendor_Name"] === vendorVal);
        }

        // Update all other filters based on this vendor data
        populateDependentFilters(relevantData);

        // Apply filters (which will now use the new options, and 'All' for reset ones)
        applyFilters();
    }

    function updateDTOptions() {
        const feederVal = document.getElementById('feederFilter').value;
        const dtSelect = document.getElementById('dtFilter');
        const currentDT = dtSelect.value;

        // Respect Vendor Context
        const vendorVal = document.getElementById('vendorFilter').value;
        let contextData = globalData;
        if (vendorVal !== 'All') {
            contextData = globalData.filter(item => item["Vendor_Name"] === vendorVal);
        }

        // Get relevant data based on Feeder selection within Vendor Context
        let relevantData = contextData;
        if (feederVal !== 'All') {
            relevantData = contextData.filter(item => item["Feeder"] === feederVal);
        }

        // Get unique DTs
        const dts = [...new Set(relevantData.map(item => item["DT Name"]))].filter(Boolean).sort();

        // Clear and populate
        dtSelect.innerHTML = '<option value="All">All DTs</option>';
        dts.forEach(dt => {
            const opt = document.createElement('option');
            opt.value = dt;
            opt.textContent = dt;
            dtSelect.appendChild(opt);
        });

        // Restore selection if valid
        if (dts.includes(currentDT)) {
            dtSelect.value = currentDT;
        } else {
            dtSelect.value = 'All';
        }

        // Trigger Upriser update
        updateUpriserOptions();
    }

    function updateUpriserOptions() {
        const dtVal = document.getElementById('dtFilter').value;
        const upriserSelect = document.getElementById('upriserFilter');
        const currentUpriser = upriserSelect.value;
        const feederVal = document.getElementById('feederFilter').value;

        // Respect Vendor Context
        const vendorVal = document.getElementById('vendorFilter').value;
        let contextData = globalData;
        if (vendorVal !== 'All') {
            contextData = globalData.filter(item => item["Vendor_Name"] === vendorVal);
        }

        // Get relevant data based on DT (and implicitly Feeder)
        let relevantData = contextData;

        if (dtVal !== 'All') {
            relevantData = contextData.filter(item => item["DT Name"] === dtVal);
        } else if (feederVal !== 'All') {
            relevantData = contextData.filter(item => item["Feeder"] === feederVal);
        }

        // Get unique Uprisers
        const uprisers = [...new Set(relevantData.map(item => item["UpriserNo"]))].filter(Boolean).sort((a, b) => a - b);

        // Clear and populate
        upriserSelect.innerHTML = '<option value="All">All Uprisers</option>';
        uprisers.forEach(upriser => {
            const opt = document.createElement('option');
            opt.value = upriser;
            opt.textContent = upriser;
            upriserSelect.appendChild(opt);
        });

        if (uprisers.some(u => String(u) === currentUpriser)) {
            upriserSelect.value = currentUpriser;
        } else {
            upriserSelect.value = 'All';
        }
    }

    function applyFilters() {
        const vendorVal = document.getElementById('vendorFilter').value;
        const buVal = document.getElementById('buFilter').value;
        const utVal = document.getElementById('utFilter').value;
        const userVal = document.getElementById('userFilter').value;
        const dtVal = document.getElementById('dtFilter').value;
        const upriserVal = document.getElementById('upriserFilter').value;
        const feederVal = document.getElementById('feederFilter').value;
        const matVal = document.getElementById('materialFilter').value;
        const dateVal = document.getElementById('dateFilter').value;

        filteredData = globalData.filter(item => {
            // Material check
            const mat = String(item["Pole Material"] || item["Material"] || item["Pole_Material"] || "").toLowerCase();
            const poleType = String(item["Type of Pole"] || "").toLowerCase();

            const matMatch = (matVal === "" ||
                (matVal === "Concrete" && (mat.includes('concrete') || poleType.includes('concrete'))) ||
                (matVal === "Wood" && (mat.includes('wood') || poleType.includes('wood')))
            );

            return (vendorVal === 'All' || item["Vendor_Name"] === vendorVal) &&
                (buVal === 'All' || item["Bussines Unit"] === buVal) &&
                (utVal === 'All' || item["Undertaking"] === utVal) &&
                (userVal === 'All' || item["User"] === userVal) &&
                (dtVal === 'All' || item["DT Name"] === dtVal) &&
                (upriserVal === 'All' || String(item["UpriserNo"]) == upriserVal) &&
                (feederVal === 'All' || item["Feeder"] === feederVal) &&
                (dateVal === 'All' || (item["Date/timestamp"] && item["Date/timestamp"].startsWith(dateVal))) &&
                matMatch;
        });

        updateDashboard();
    }

    function updateDashboard() {
        const fieldCharts = document.getElementById('charts');
        const varianceCharts = document.getElementById('variance-charts');

        if (viewMode === 'boq') {
            // Show Variance View
            if (fieldCharts) fieldCharts.classList.add('hidden');
            if (varianceCharts) varianceCharts.classList.remove('hidden');
            updateKPIs(); // Will handle variance logic
            renderVarianceCharts();
            renderDTTable(); // Will handle variance columns
        } else {
            // Show Field View
            if (fieldCharts) fieldCharts.classList.remove('hidden');
            if (varianceCharts) varianceCharts.classList.add('hidden');
            updateKPIs();
            renderUserPerformanceChart();
            renderProjectVelocityChart();
            renderPoleTypeChart();
            renderStaffIssuesChart();
            renderUndertakingChart();
            renderFeederChart();
            renderVendorPerformanceCharts();
            renderDTTable();
        }
        // Map is shared or hidden? User didn't specify. Left as is (always showing map based on field data).
        // Maybe hide map in variance mode? User said "View Mode: Field Captures Only | BOQ vs. Actual".
        // Usually map is useful. Detailed request didn't say hide map.
        renderMap();
        updateKeyInsights();
        renderStrategicRecommendations();
    }


    function updateKPIs() {
        // Helper to formatting numbers
        const fmt = n => n ? n.toLocaleString() : '0';

        // 1. Calculate Metrics

        // --- A. Records (Poles) ---
        const boqRecords = (viewMode === 'boq' && boqData.length)
            ? boqData.reduce((sum, d) => sum + (parseInt(d["POLES Grand Total"]) || 0), 0)
            : 0;
        const actRecords = filteredData.length;
        updateModernCard('records', boqRecords, actRecords);

        // --- B. Good Poles (Concrete/Good) ---
        const boqGood = (viewMode === 'boq' && boqData.length)
            ? boqData.reduce((sum, d) => sum + (parseInt(d["GOOD"]) || 0), 0)
            : 0;
        const actGood = filteredData.filter(d => (d.Issue_Type === 'Good Condition')).length;
        updateModernCard('concrete', boqGood, actGood);

        // --- C. Bad Poles (Wooden/Replace) ---
        const boqBad = (viewMode === 'boq' && boqData.length)
            ? boqData.reduce((sum, d) => sum + (parseInt(d["BAD"]) || 0), 0)
            : 0;
        const actBad = filteredData.filter(d => (d.Issue_Type !== 'Good Condition')).length;
        updateModernCard('wooden', boqBad, actBad);

        // --- D. New Poles (Install) ---
        const boqNew = (viewMode === 'boq' && boqData.length)
            ? boqData.reduce((sum, d) => sum + (parseInt(d["NEW POLE"]) || 0), 0)
            : 0;
        // Logic for Actual New Poles: Check Pole_Type or Issue_Type for 'New'
        // If not found, default to 0 to avoid misleading data
        const actNew = filteredData.filter(d =>
            (d.Pole_Type && d.Pole_Type.toLowerCase().includes('new')) ||
            (d.Issue_Type && d.Issue_Type.toLowerCase().includes('new'))
        ).length;
        updateModernCard('users', boqNew, actNew);

        // --- E. Feeders ---
        const boqFeeders = (viewMode === 'boq' && boqData.length) ? new Set(boqData.map(d => d["FEEDER NAME"])).size : 0;
        const actFeeders = new Set(filteredData.map(d => d.Feeder)).size;
        updateModernCard('feeders', boqFeeders, actFeeders);

        // --- F. DTs ---
        const boqDTs = (viewMode === 'boq' && boqData.length) ? new Set(boqData.map(d => d["DT NAME"])).size : 0;
        const actDTs = new Set(filteredData.map(d => d["DT Name"] || d["DT_Name"])).size;
        updateModernCard('dts', boqDTs, actDTs);

        // --- G. Buildings ---
        // BOQ for buildings might not exist, defaulting to 0 for now.
        const boqBuildings = 0;
        const actBuildings = filteredData.reduce((sum, item) => sum + (parseInt(item["No of Buildings Connected to the Pole"]) || 0), 0);
        updateModernCard('buildings', boqBuildings, actBuildings);
    }

    function updateModernCard(suffix, boqVal, actVal) {
        const elBoq = document.getElementById(`kpi-boq-${suffix}`);
        const elAct = document.getElementById(`kpi-act-${suffix}`);
        const elProg = document.getElementById(`kpi-prog-${suffix}`);
        const elBar = document.getElementById(`kpi-bar-${suffix}`);
        const elRem = document.getElementById(`kpi-rem-${suffix}`);

        if (!elAct) return;

        // Set Values
        if (elBoq) elBoq.textContent = (viewMode === 'boq' && boqData.length) ? boqVal.toLocaleString() : '-';
        elAct.textContent = actVal.toLocaleString();

        // Calculate Progress
        let pct = 0;
        if (viewMode === 'boq' && boqVal > 0) {
            pct = (actVal / boqVal) * 100;
        }

        const displayPct = pct.toFixed(1) + '%';
        const barWidth = Math.min(pct, 100) + '%';

        if (elProg) elProg.textContent = displayPct;
        if (elBar) elBar.style.width = barWidth;

        // Remaining
        if (elRem) {
            if (viewMode === 'boq' && boqVal > 0) {
                const rem = boqVal - actVal;
                elRem.textContent = `Remaining: ${Math.max(0, rem).toLocaleString()}`;
            } else {
                elRem.textContent = 'Remaining: -';
            }
        }
        if (suffix === 'records') {
            const completionText = document.getElementById('system-completion-text');
            const completionBar = document.getElementById('system-completion-bar');
            if (completionText && completionBar) {
                // If viewMode is field (default), boqVal is 0, so completion is 0%.
                // We likely want global BOQ vs global Actual for "Project Completion" regardless of view mode?
                // Visual indicates "4.9%". Let's use global BOQ sum if possible.
                // Re-calculate global progress if needed, or use current KPI logic.
                // KPI logic uses boqVal which is 0 in 'field' mode.
                // Let's force calculate global BOQ for this status bar if boqVal is 0.

                let totalBoq = boqVal;
                if (totalBoq === 0 && boqData.length > 0) {
                    totalBoq = boqData.reduce((sum, d) => sum + (parseInt(d["POLES Grand Total"]) || 0), 0);
                }

                let systemPct = 0;
                if (totalBoq > 0) {
                    systemPct = (actVal / totalBoq) * 100;
                }

                completionText.textContent = systemPct.toFixed(1) + '%';
                completionBar.style.width = Math.min(systemPct, 100) + '%';
            }
        }
    }

    // --- Chart Rendering Functions ---

    // 1. User Performance (Bar Chart)
    function renderUserPerformanceChart() {
        const userCounts = {};
        const userVendors = {};

        filteredData.forEach(d => {
            userCounts[d.User] = (userCounts[d.User] || 0) + 1;
            if (!userVendors[d.User]) userVendors[d.User] = d.Vendor_Name;
        });

        const sortedUsers = Object.entries(userCounts).sort((a, b) => b[1] - a[1]);
        const xUsernames = sortedUsers.map(u => u[0]);
        // Map usernames to full names, fallback to username if not found
        const xLabels = xUsernames.map(u => userFullNames[u] || u);
        const y = sortedUsers.map(u => u[1]);

        // Assign colors based on vendor
        const colors = xUsernames.map(user => {
            const vendor = userVendors[user];
            if (vendor === 'ETC Workforce') return '#0EA5E9'; // Blue
            if (vendor === 'Jesom Technology') return '#f97316'; // Orange
            return '#a0a0a0'; // Grey for others
        });

        const trace = {
            x: xLabels, // Use full names here
            y: y,
            type: 'bar',
            marker: {
                color: colors
            }
        };

        const layout = {
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            font: { color: '#fafafa' },
            margin: { t: 30, b: 120, l: 50, r: 20 }, // INCREASE BOTTOM MARGIN for names
            xaxis: { title: '', tickangle: -45 },
            yaxis: { title: 'Records Captured' },
            // Add a manual legend since we are using a single trace with multiple colors
            annotations: [
                {
                    xref: 'paper', yref: 'paper',
                    x: 1, y: 1,
                    xanchor: 'right', yanchor: 'bottom',
                    text: '<span style="color:#0EA5E9">■</span> ETC Workforce  <span style="color:#f97316">■</span> Jesom Technology',
                    showarrow: false,
                    font: { size: 12, color: '#fafafa' }
                }
            ]
        };

        Plotly.newPlot('userPerformanceChart', [trace], layout, { responsive: true });
    }

    // 2. Project Velocity (Area Chart Comparison)
    function renderProjectVelocityChart() {
        // Group by Date and Vendor
        const dateVendorCounts = {}; // { "date": { "ETC": count, "Jesom": count } }

        filteredData.forEach(d => {
            const date = d["Date/timestamp"].split(' ')[0]; // Extract Date part
            const vendor = d.Vendor_Name;

            if (!dateVendorCounts[date]) {
                dateVendorCounts[date] = { 'ETC Workforce': 0, 'Jesom Technology': 0, 'Other': 0 };
            }
            if (dateVendorCounts[date][vendor] !== undefined) {
                dateVendorCounts[date][vendor]++;
            } else {
                // Try to be smart about 'Other' merging into one if needed, but for now allow separate
                dateVendorCounts[date]['Other']++;
            }
        });

        // Sort dates
        const sortedDates = Object.keys(dateVendorCounts).sort((a, b) => new Date(a) - new Date(b));

        const yETC = sortedDates.map(date => dateVendorCounts[date]['ETC Workforce']);
        const yJesom = sortedDates.map(date => dateVendorCounts[date]['Jesom Technology']);

        const traceETC = {
            x: sortedDates,
            y: yETC,
            name: 'ETC Workforce (Blue)',
            type: 'scatter',
            mode: 'lines+markers+text',
            text: yETC.map(String),
            textposition: 'top center',
            fill: 'tozeroy',
            line: { color: '#0EA5E9' }, // Blue
            marker: { size: 6 }
        };

        const traceJesom = {
            x: sortedDates,
            y: yJesom,
            name: 'Jesom Technology (Orange)',
            type: 'scatter',
            mode: 'lines+markers+text',
            text: yJesom.map(String),
            textposition: 'top center',
            fill: 'tozeroy',
            line: { color: '#f97316' }, // Orange
            marker: { size: 6 }
        };

        const layout = {
            background_color: '#333333', // Dark background for chart area per image appearance (optional but sticking to theme first)
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(255,255,255,0.02)', // Slightly lighter plot area
            font: { color: '#fafafa' },
            xaxis: {
                title: '',
                tickformat: '%b %d, %Y', // e.g. Jan 30, 2026
                tickvals: sortedDates // Ensure all dates are shown if needed, or let Plotly handle it
            },
            yaxis: { title: '' },
            margin: { t: 40, l: 40, r: 20, b: 40 },
            showlegend: true,
            legend: { orientation: "h", y: -0.2 }
        };

        Plotly.newPlot('projectVelocityChart', [traceETC, traceJesom], layout, { responsive: true });
    }

    // 3. Pole Type Distribution (highcharts 3D Pie Chart)
    function renderPoleTypeChart() {
        const counts = {};
        filteredData.forEach(d => {
            const type = d["Type of Pole"] || "Unknown";
            counts[type] = (counts[type] || 0) + 1;
        });

        const data = Object.keys(counts).map(key => {
            let color = '#a0a0a0';
            const upper = key.toUpperCase();
            if (upper.includes('CONCRETE')) color = '#10b981';
            if (upper.includes('WOOD')) color = '#ef4444';

            return {
                name: key,
                y: counts[key],
                color: color
            };
        });

        Highcharts.chart('poleTypeChart', {
            chart: {
                type: 'pie',
                backgroundColor: 'rgba(0,0,0,0)',
                options3d: {
                    enabled: true,
                    alpha: 45
                }
            },
            title: {
                text: null
            },
            tooltip: {
                pointFormat: '{series.name}: <b>{point.percentage:.1f}%</b>'
            },
            plotOptions: {
                pie: {
                    innerSize: '50%', // Create Doughnut
                    size: '60%',      // Reduce visual width
                    depth: 35,
                    allowPointSelect: true,
                    cursor: 'pointer',
                    dataLabels: {
                        enabled: true,
                        format: '<b>{point.name}</b>: {point.y} ({point.percentage:.1f} %)',
                        style: {
                            color: '#e4e5e7',
                            textOutline: 'none',
                            fontSize: '8px'
                        },
                        connectorColor: 'silver',
                        softConnector: true
                    },
                    showInLegend: true
                }
            },
            legend: {
                itemStyle: {
                    color: '#e4e5e7',
                    fontWeight: 'normal',
                    fontSize: '9px'
                },
                itemHoverStyle: {
                    color: '#ffffff'
                }
            },
            series: [{
                name: 'Distribution',
                data: data
            }],
            credits: {
                enabled: false
            }
        });
    }

    // 3.5 Issues by Staff (Stacked Bar)
    function renderStaffIssuesChart() {
        // Group by User -> Issue Type -> Count
        const userIssues = {};
        const issuesSet = new Set();

        filteredData.forEach(d => {
            const user = d.User;
            const issue = d.Issue_Type;
            if (issue === 'Good Condition') return; // Filter out 'Good' to focus on issues? Or keep all? Prompt implies distinct issues. Let's filter 'Good' to make it look like the example "Reported Issues".

            issuesSet.add(issue);
            if (!userIssues[user]) userIssues[user] = {};
            userIssues[user][issue] = (userIssues[user][issue] || 0) + 1;
        });

        const issueTypes = Array.from(issuesSet); // e.g. Broken, Crooked...

        // Sort users by total issues
        const sortedUsers = Object.keys(userIssues).sort((a, b) => {
            const totalA = Object.values(userIssues[a]).reduce((s, c) => s + c, 0);
            const totalB = Object.values(userIssues[b]).reduce((s, c) => s + c, 0);
            return totalB - totalA;
        });

        // Prepare Traces (one per issue type)
        const traces = issueTypes.map(issue => {
            return {
                x: sortedUsers.map(u => userFullNames[u] || u),
                y: sortedUsers.map(u => userIssues[u][issue] || 0),
                name: issue,
                type: 'bar'
            };
        });

        const layout = {
            barmode: 'stack',
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            font: { color: '#e4e5e7' },
            xaxis: { title: '', tickangle: -45 },
            yaxis: { title: 'Number of Issues' },
            margin: { t: 30, b: 100, l: 50, r: 20 },
            legend: { orientation: 'h', y: 1.1 }
        };

        Plotly.newPlot('staffIssuesChart', traces, layout, { responsive: true });
    }

    // 4. Undertaking Breakdown (Bar Chart - Horizontal)
    function renderUndertakingChart() {
        const counts = {};
        filteredData.forEach(d => {
            counts[d["Undertaking"]] = (counts[d["Undertaking"]] || 0) + 1;
        });

        const sorted = Object.entries(counts).sort((a, b) => a[1] - b[1]); // Ascending for horizontal bar
        const y = sorted.map(i => i[0]);
        const x = sorted.map(i => i[1]);

        const trace = {
            x: x,
            y: y,
            type: 'bar',
            orientation: 'h',
            marker: {
                color: '#f59e0b'
            }
        };

        const layout = {
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            font: { color: '#e4e5e7' },
            margin: { t: 20, l: 100, r: 20, b: 40 },
            xaxis: { title: 'Count' }
        };

        Plotly.newPlot('undertakingChart', [trace], layout, { responsive: true });
    }

    // 5. Vendor Performance Comparison (Total Records & Run Rate)
    function renderVendorPerformanceCharts() {
        // Track records and unique (User + Date) combinations for Man-Days
        const vendorData = {
            'ETC Workforce': { records: 0, manDays: new Set() },
            'Jesom Technology': { records: 0, manDays: new Set() }
        };

        filteredData.forEach(d => {
            const vendor = d.Vendor_Name;
            const date = d["Date/timestamp"].split(' ')[0];
            const user = d.User;

            if (vendorData[vendor]) {
                vendorData[vendor].records++;
                vendorData[vendor].manDays.add(`${user}|${date}`); // Unique Man-Day
            }
        });

        const vendors = ['ETC Workforce', 'Jesom Technology'];

        // Data for Chart 1: Total Records
        const totalRecords = vendors.map(v => vendorData[v].records);

        // Data for Chart 2: Avg Run Rate per Field Officer (Records / Man-Days)
        const runRates = vendors.map(v => {
            const days = vendorData[v].manDays.size || 1;
            return (vendorData[v].records / days);
        });

        const blueColor = '#0EA5E9'; // e.g. bright blue
        const redColor = '#f97316'; // Jesom Orange (formerly red)

        // --- Chart 1: Total Records ---
        const traceTotal = {
            x: vendors,
            y: totalRecords,
            type: 'bar',
            text: totalRecords.map(String),
            textposition: 'auto',
            marker: {
                color: [blueColor, redColor]
            }
        };

        const layoutTotal = {
            title: null, // Title in HTML now
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            font: { color: '#fafafa' },
            xaxis: { title: '' },
            yaxis: { title: '', showgrid: true, gridcolor: '#334155' },
            margin: { t: 20, b: 40, l: 40, r: 20 },
            height: 350 // Adjusted for single card
        };

        Plotly.newPlot('vendorTotalChart', [traceTotal], layoutTotal, { responsive: true });

        // --- Chart 2: Run Rate ---
        const traceRunRate = {
            x: vendors,
            y: runRates,
            type: 'bar',
            text: runRates.map(v => v.toFixed(1)),
            textposition: 'auto',
            marker: {
                color: [blueColor, redColor]
            },
            name: 'Run Rate'
        };

        // Target Line (50/day)
        const targetLine = {
            type: 'line',
            x0: -0.5,
            x1: 1.5,
            y0: 50,
            y1: 50,
            line: {
                color: '#10b981', // green
                width: 2,
                dash: 'dash'
            }
        };

        const layoutRunRate = {
            title: null, // Title in HTML now
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            font: { color: '#fafafa' },
            xaxis: { title: '' },
            yaxis: { title: '', showgrid: true, gridcolor: '#334155', range: [0, Math.max(60, Math.max(...runRates) * 1.1)] }, // Ensure grid scale fits target line
            margin: { t: 20, b: 40, l: 40, r: 20 },
            height: 350, // Adjusted for single card
            shapes: [targetLine],
            annotations: [{
                x: 1,
                y: 52,
                xref: 'x',
                yref: 'y',
                text: 'Target: 50/day',
                showarrow: false,
                font: { color: '#10b981' }
            }]
        };

        Plotly.newPlot('vendorRunRateChart', [traceRunRate], layoutRunRate, { responsive: true });
    }


    // --- Column Definitions ---
    const tableColumns = [
        { id: 'col-index', label: '#', visible: true },
        { id: 'col-dtName', label: 'DT Name', visible: true },
        { id: 'col-feeder', label: 'Feeder Name', visible: true },
        { id: 'col-bu', label: 'BU', visible: true },
        { id: 'col-undertaking', label: 'Undertaking', visible: true },
        { id: 'col-vendor', label: 'Vendor', visible: true },
        { id: 'col-users', label: 'Field Officers', visible: true },
        { id: 'col-boqTotal', label: 'Total (BOQ)', visible: true },
        { id: 'col-actualTotal', label: 'Actual', visible: true },
        { id: 'col-remaining', label: 'Remaining', visible: true },
        { id: 'col-concrete', label: 'Concrete', visible: true },
        { id: 'col-wooden', label: 'Wooden', visible: true },
        { id: 'col-progress', label: 'Progress', visible: true },
        { id: 'col-status', label: 'Status', visible: true }
    ];

    function initColumnFilters() {
        const menu = document.getElementById('columnFilterMenu');
        const btn = document.getElementById('columnFilterBtn');
        if (!menu || !btn) return;

        // Toggle Menu
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            menu.style.display = menu.style.display === 'flex' ? 'none' : 'flex';
        });

        // Close on outside click
        document.addEventListener('click', (e) => {
            if (!menu.contains(e.target) && !btn.contains(e.target)) {
                menu.style.display = 'none';
            }
        });

        // Populate
        menu.innerHTML = '';
        tableColumns.forEach((col) => {
            const label = document.createElement('label');
            label.className = 'col-check-item';

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.checked = col.visible;
            checkbox.addEventListener('change', () => {
                col.visible = checkbox.checked;
                updateColumnVisibility();
            });

            label.appendChild(checkbox);
            label.appendChild(document.createTextNode(col.label));
            menu.appendChild(label);
        });

        // Initial Visibility Set
        updateColumnVisibility();
    }

    function updateColumnVisibility() {
        // Update Header
        tableColumns.forEach(col => {
            const th = document.querySelector(`#dtTable th.${col.id}`);
            if (th) th.style.display = col.visible ? '' : 'none';
        });

        // Update Body (by re-rendering)
        renderDTTable();
    }

    function initSearchSuggestions() {
        const input = document.getElementById('dtSearchInput');
        const list = document.getElementById('searchSuggestions');
        if (!input || !list) return;

        input.addEventListener('input', () => {
            const val = input.value.trim().toLowerCase();

            // Trigger render on every keystroke
            renderDTTable();

            if (val.length < 2) {
                list.style.display = 'none';
                return;
            }

            // Get unique suggestions
            const suggestions = new Set();
            const data = getEnhancedDTData();

            for (const item of data) {
                if (suggestions.size >= 8) break; // Limit suggestions

                const candidates = [
                    item.dtName,
                    item.vendor,
                    item.feeder,
                    item.bu,
                    item.undertaking
                ];

                candidates.forEach(c => {
                    if (c && String(c).toLowerCase().includes(val)) {
                        suggestions.add(String(c));
                    }
                });

                // Users
                item.users.forEach(u => {
                    const name = userFullNames[u] || u;
                    if (name && name.toLowerCase().includes(val)) suggestions.add(name);
                });
            }

            // Render
            const results = Array.from(suggestions).slice(0, 8);
            if (results.length > 0) {
                list.innerHTML = results.map(s =>
                    `<div style="padding: 8px 12px; cursor: pointer; color: #e2e8f0; border-bottom: 1px solid #334155;">${s}</div>`
                ).join('');

                list.style.display = 'flex';

                // Click handler
                Array.from(list.children).forEach(div => {
                    div.addEventListener('click', () => {
                        input.value = div.textContent;
                        list.style.display = 'none';
                        renderDTTable();
                    });

                    // Hover effect
                    div.addEventListener('mouseenter', () => div.style.backgroundColor = '#334155');
                    div.addEventListener('mouseleave', () => div.style.backgroundColor = 'transparent');
                });
            } else {
                list.style.display = 'none';
            }
        });

        // Hide on click outside
        document.addEventListener('click', (e) => {
            if (!input.contains(e.target) && !list.contains(e.target)) {
                list.style.display = 'none';
            }
        });
    }

    // 6. Detailed DT Analysis Table (Enhanced)
    function renderDTTable() {
        const tbody = document.querySelector('#dtTable tbody');
        if (!tbody) return;

        tbody.innerHTML = '';
        const searchVal = (document.getElementById('dtSearchInput')?.value || '').toLowerCase();

        // 1. Get Enhanced Data (Union of BOQ and Field)
        const data = getEnhancedDTData();

        // 2. Filter by Search Input (Interactive)
        const filtered = data.filter(item => {
            if (!searchVal) return true;
            return (
                (item.dtName || '').toLowerCase().includes(searchVal) ||
                (item.vendor || '').toLowerCase().includes(searchVal) ||
                (item.feeder || '').toLowerCase().includes(searchVal) ||
                (item.bu || '').toLowerCase().includes(searchVal) ||
                (item.undertaking || '').toLowerCase().includes(searchVal) ||
                item.users.some(u => String(userFullNames[u] || u || '').toLowerCase().includes(searchVal))
            );
        });

        // 3. Update Info Count
        const infoEl = document.getElementById('tableInfo');
        if (infoEl) infoEl.textContent = `Showing ${filtered.length} of ${data.length} DTs`;

        // 4. Pagination Logic
        const totalItems = filtered.length;
        const totalPages = Math.ceil(totalItems / rowsPerPage);

        // Adjust currentPage if out of bounds
        if (currentPage > totalPages && totalPages > 0) currentPage = totalPages;
        if (currentPage < 1 && totalPages > 0) currentPage = 1; // Should happen?
        if (totalPages === 0) currentPage = 1; // If no items, reset to page 1

        const startIndex = (currentPage - 1) * rowsPerPage;
        const endIndex = startIndex + rowsPerPage;
        const paginatedData = filtered.slice(startIndex, endIndex);

        // Helper to check visibility
        const isVisible = (id) => {
            const col = tableColumns.find(c => c.id === id);
            return col ? col.visible : true;
        };

        // 5. Render Rows
        paginatedData.forEach((row, index) => {
            const tr = document.createElement('tr');

            // Vendor Tag
            let vendorClass = '';
            if (row.vendor === 'ETC Workforce') vendorClass = 'vendor-etc';
            if (row.vendor === 'Jesom Technology') vendorClass = 'vendor-jesom';

            // Progress Bar / Status Logic
            const progress = row.boqTotal > 0 ? (row.actualTotal / row.boqTotal) * 100 : 0;
            let status = 'In Progress';
            let statusColor = '#f59e0b'; // Orange

            if (row.actualTotal === 0) {
                status = 'Not Started';
                statusColor = '#ef4444'; // Red
            } else if (progress >= 100) {
                status = 'Completed';
                statusColor = '#10b981'; // Green
            } else if (progress > 90) {
                status = 'Near Completion';
                statusColor = '#3b82f6'; // Blue
            }

            // User Names
            const userNames = row.users.map(u => userFullNames[u] || u).join(', ');
            // Absolute index for numbering
            const absIndex = startIndex + index + 1;

            let rowHtml = '';

            if (isVisible('col-index')) rowHtml += `<td class="col-index" style="text-align: center;">${absIndex}</td>`;
            if (isVisible('col-dtName')) rowHtml += `<td class="col-dtName" style="font-weight: 500; color: #fff;">${row.dtName}</td>`;
            if (isVisible('col-feeder')) rowHtml += `<td class="col-feeder">${row.feeder}</td>`;
            if (isVisible('col-bu')) rowHtml += `<td class="col-bu">${row.bu}</td>`;
            if (isVisible('col-undertaking')) rowHtml += `<td class="col-undertaking">${row.undertaking}</td>`;
            if (isVisible('col-vendor')) rowHtml += `<td class="col-vendor"><span class="vendor-tag ${vendorClass}">${row.vendor}</span></td>`;
            if (isVisible('col-users')) rowHtml += `<td class="col-users" style="max-width: 140px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${userNames}">${userNames}</td>`;
            if (isVisible('col-boqTotal')) rowHtml += `<td class="col-boqTotal" style="text-align: center; font-weight: bold; color: #0EA5E9;">${row.boqTotal}</td>`;
            if (isVisible('col-actualTotal')) rowHtml += `<td class="col-actualTotal" style="text-align: center;">${row.actualTotal}</td>`;
            if (isVisible('col-remaining')) rowHtml += `<td class="col-remaining" style="text-align: center; color: #a0a0a0;">${Math.max(0, row.boqTotal - row.actualTotal)}</td>`;
            if (isVisible('col-concrete')) rowHtml += `<td class="col-concrete" style="text-align: center;">${row.concrete}</td>`;
            if (isVisible('col-wooden')) rowHtml += `<td class="col-wooden" style="text-align: center;">${row.wooden}</td>`;
            if (isVisible('col-progress')) rowHtml += `<td class="col-progress" style="width: 70px;">
                    <div style="display: flex; align-items: center; gap: 4px;">
                        <div style="flex-grow: 1; height: 4px; background: #333; border-radius: 2px; overflow: hidden;">
                            <div style="width: ${Math.min(100, progress)}%; height: 100%; background: ${statusColor};"></div>
                        </div>
                        <span style="font-size: 0.8em; color: ${statusColor};">${progress.toFixed(0)}%</span>
                    </div>
                </td>`;
            if (isVisible('col-status')) rowHtml += `<td class="col-status"><span style="font-size: 0.8em; padding: 1px 6px; border-radius: 8px; background: ${statusColor}20; color: ${statusColor}; border: 1px solid ${statusColor}40; white-space: nowrap;">${status}</span></td>`;

            tr.innerHTML = rowHtml;
            tbody.appendChild(tr);
        });

        // 6. Render Pagination Controls
        renderPaginationControls(filtered.length);
    }

    function renderPaginationControls(totalItems) {
        const container = document.getElementById('paginationControls');
        if (!container) return;
        container.innerHTML = '';

        const totalPages = Math.ceil(totalItems / rowsPerPage);
        if (totalPages <= 1) return;

        const createBtn = (text, page, isActive = false, isDisabled = false) => {
            const btn = document.createElement('button');
            btn.className = `page-btn ${isActive ? 'active' : ''}`;
            btn.textContent = text;
            if (isDisabled) btn.disabled = true;
            else {
                btn.onclick = () => {
                    currentPage = page;
                    renderDTTable();
                };
            }
            return btn;
        };

        // Prev Button
        container.appendChild(createBtn('<', currentPage - 1, false, currentPage === 1));

        // Page Range Logic (Show up to 6 pages)
        const maxVisible = 6;
        let startPage = 1;
        let endPage = Math.min(totalPages, maxVisible);

        if (currentPage > 3 && totalPages > maxVisible) {
            // Center user in the window if possible
            startPage = Math.max(1, currentPage - 2);
            endPage = Math.min(totalPages, startPage + maxVisible - 1);

            // Adjust start if end is capped
            if (endPage === totalPages) {
                startPage = Math.max(1, endPage - maxVisible + 1);
            }
        }

        for (let i = startPage; i <= endPage; i++) {
            container.appendChild(createBtn(i, i, i === currentPage));
        }

        // Next Button
        container.appendChild(createBtn('>', currentPage + 1, false, currentPage === totalPages));
    }

    function resetFilters() {
        // 1. Reset View Mode first
        viewMode = 'field';
        const toggle = document.getElementById('viewModeToggle');
        if (toggle) toggle.checked = false;

        // 2. Clear Search Input
        const searchInput = document.getElementById('dtSearchInput');
        if (searchInput) searchInput.value = '';

        // 3. Reset Pagination
        currentPage = 1;

        // 4. Reset Filters UI & Data
        // Re-populate from scratch (this resets options to global state)
        populateFilters();

        // Ensure all selects are set to 'All' (populateFilters might do this implicitly, but let's be sure)
        const filterIds = [
            'vendorFilter', 'buFilter', 'utFilter', 'userFilter',
            'feederFilter', 'dtFilter', 'upriserFilter', 'materialFilter', 'dateFilter'
        ];

        filterIds.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = 'All';
        });

        // 5. Update Dashboard (This will rebuild filteredData from globalData based on the 'All' selections)
        applyFilters();
    }

    function getEnhancedDTData() {
        const map = {};

        // 1. Process Field Data
        filteredData.forEach(d => {
            const dtName = (d["DT Name"] || "Unknown DT").trim();
            const feeder = (d["Feeder"] || "Unknown Feeder").trim();
            const key = `${feeder}|${dtName}`.toUpperCase();

            if (!map[key]) {
                map[key] = {
                    key,
                    dtName,
                    feeder,
                    bu: d["Bussines Unit"] || "-",
                    undertaking: d["Undertaking"] || "-",
                    vendor: d["Vendor_Name"] || "-",
                    users: new Set(),
                    boqTotal: 0, // Will fill from BOQ
                    actualTotal: 0,
                    concrete: 0,
                    wooden: 0
                };
            }

            map[key].actualTotal++;
            map[key].users.add(d.User);

            // Material
            const mat = String(d["Pole Material"] || d["Material"] || d["Pole_Material"] || "").toLowerCase();
            const type = String(d["Type of Pole"] || "").toLowerCase();
            if (mat.includes('concrete') || type.includes('concrete')) map[key].concrete++;
            if (mat.includes('wood') || type.includes('wood')) map[key].wooden++;
        });

        // 2. Process BOQ Data (Fill Targets)
        // Respect Feeder/DT filters if possible, but for "Total (BOQ)", usually we want the Static BOQ target for that DT.
        // However, we should filter BOQ by the global dashboard filters TO AN EXTENT (e.g. if I selected a Feeder, I only want DTs in that Feeder).
        // `filteredData` is already filtered. `boqData` is just raw.
        // I need to iterate `boqData` and match. 
        // Also if a DT is in BOQ but NOT in field data, we should add it?
        // Yes, to show "0 Actual" and "Status: Not Started".

        // Apply same filters to BOQ as Dashboard?
        // The dashboard filters (bu, ut, vendor...) apply to Field Data.
        // BOQ only has Feeder/DT.
        // If I filter by Vendor=ETC, I should only see DTs assigned to ETC?
        // But BOQ doesn't have Vendor.
        // Only Field Data determines Vendor.
        // So if I filter by Vendor, I implicitly filter out "Not Started" DTs because they have no Vendor assigned in Field Data yet?
        // UNLESS we have a mapping of BOQ DTs to Vendors. We don't.
        // So: If filtered by Vendor, we only show DTs that have started (have field data).
        // If NO Vendor filter (All), we show everything.
        // This suggests:
        // - Iterate field map (which respects all filters).
        // - Iterate BOQ. If BOQ item matches a key in field map, update boqTotal.
        // - IF BOQ item does NOT match field map:
        //   - IF "All" filters are selected (or at least Vendor is All), add it as "Not Started".
        //   - IF filters are active (e.g. Vendor=ETC), do NOT add it (because we don't know if it belongs to ETC).

        const vendorFilter = document.getElementById('vendorFilter')?.value || 'All';
        const isFiltered = vendorFilter !== 'All'; // Simplified check. Just checking Vendor.

        boqData.forEach(d => {
            const dtName = (d["DT NAME"] || "Unknown DT").trim();
            const feeder = (d["FEEDER NAME"] || "Unknown Feeder").trim();
            const key = `${feeder}|${dtName}`.toUpperCase();

            // Check filters (Feeder/DT)
            const selFeeder = document.getElementById('feederFilter')?.value;
            const selDT = document.getElementById('dtFilter')?.value;

            if (selFeeder && selFeeder !== 'All' && feeder !== selFeeder) return;
            if (selDT && selDT !== 'All' && dtName !== selDT) return;


            if (map[key]) {
                // Exists in field data (so it passed field filters)
                map[key].boqTotal += (parseInt(d["POLES Grand Total"]) || 0);
            } else {
                // Not in field data.
                // Only add if we are not strictly filtering by attributes we determine from field (like Vendor, User, Material, BU, UT).
                // If I filtered by "Concrete", I can't show a BOQ-only item because I don't know if it will be concrete.
                // So, if ANY filter (other than Feeder/DT) is active, we might skip BOQ-only items to avoid showing unrelated data?
                // OR we just show them as "No Data".
                // But the user request implies a management dashboard.
                // Let's safe side: Only add BOQ-only items if NO major field-dependent filters are active.
                // Major filters: Vendor, BU, Undertaking, User, Material.

                // Active Filters Check
                const fVendor = document.getElementById('vendorFilter').value;
                const fBU = document.getElementById('buFilter').value;
                const fUT = document.getElementById('utFilter').value;
                const fUser = document.getElementById('userFilter').value;
                const fMat = document.getElementById('materialFilter')?.value || '';

                const hasFieldFilter = fVendor !== 'All' || fBU !== 'All' || fUT !== 'All' || fUser !== 'All' || fMat !== '';

                if (!hasFieldFilter) {
                    map[key] = {
                        key,
                        dtName,
                        feeder,
                        bu: "-",
                        undertaking: "-",
                        vendor: "Pending", // No vendor assigned yet
                        users: [],
                        boqTotal: (parseInt(d["POLES Grand Total"]) || 0),
                        actualTotal: 0,
                        concrete: 0,
                        wooden: 0
                    };
                }
            }
        });

        // 3. Convert Map to Array and Finalize
        return Object.values(map).map(item => ({
            ...item,
            users: Array.from(item.users)
        }));
    }

    // 6. Detailed DT Analysis Table (Enhanced)
    function renderDTTable() {
        const tbody = document.querySelector('#dtTable tbody');
        if (!tbody) return;

        tbody.innerHTML = '';
        const searchVal = (document.getElementById('dtSearchInput')?.value || '').toLowerCase();

        // 1. Get Enhanced Data (Union of BOQ and Field)
        const data = getEnhancedDTData();

        // 2. Filter by Search Input
        const filtered = data.filter(item => {
            if (!searchVal) return true;
            return (
                (item.dtName || '').toLowerCase().includes(searchVal) ||
                (item.vendor || '').toLowerCase().includes(searchVal) ||
                item.users.some(u => String(userFullNames[u] || u || '').toLowerCase().includes(searchVal))
            );
        });

        // 3. Update Info Count
        const infoEl = document.getElementById('tableInfo');
        if (infoEl) infoEl.textContent = `Showing ${filtered.length} of ${data.length} DTs`;

        // 4. Render Rows
        // 4. Pagination Logic
        const totalRows = filtered.length;
        const totalPages = Math.ceil(totalRows / rowsPerPage);

        // Ensure currentPage is valid
        if (currentPage > totalPages) currentPage = Math.max(1, totalPages);

        const startIndex = (currentPage - 1) * rowsPerPage;
        const endIndex = startIndex + rowsPerPage;

        const paginatedRows = filtered.slice(startIndex, endIndex);

        // 4b. Render Rows
        paginatedRows.forEach((row, index) => {
            const tr = document.createElement('tr');
            const globalIndex = startIndex + index + 1;

            // Vendor Tag
            let vendorClass = '';
            if (row.vendor === 'ETC Workforce') vendorClass = 'vendor-etc';
            if (row.vendor === 'Jesom Technology') vendorClass = 'vendor-jesom';

            // Progress Bar / Status Logic
            const progress = row.boqTotal > 0 ? (row.actualTotal / row.boqTotal) * 100 : 0;
            let status = 'In Progress';
            let statusColor = '#f59e0b'; // Orange

            if (row.actualTotal === 0) {
                status = 'Not Started';
                statusColor = '#ef4444'; // Red
            } else if (progress >= 100) {
                status = 'Completed';
                statusColor = '#10b981'; // Green
            } else if (progress > 90) {
                status = 'Near Completion';
                statusColor = '#3b82f6'; // Blue
            }

            // User Names
            const userNames = row.users.map(u => userFullNames[u] || u).join(', ');

            // Add classes for column visibility
            tr.innerHTML = `
                <td class="col-index" style="text-align: center;">${globalIndex}</td>
                <td class="col-dtName" style="font-weight: 500; color: #fff;">${row.dtName}</td>
                <td class="col-feeder">${row.feeder}</td>
                <td class="col-bu">${row.bu}</td>
                <td class="col-undertaking">${row.undertaking}</td>
                <td class="col-vendor"><span class="vendor-tag ${vendorClass}">${row.vendor}</span></td>
                <td class="col-users" style="max-width: 140px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${userNames}">${userNames}</td>
                <td class="col-boqTotal" style="text-align: center; font-weight: bold; color: #0EA5E9;">${row.boqTotal}</td>
                <td class="col-actualTotal" style="text-align: center;">${row.actualTotal}</td>
                <td class="col-remaining" style="text-align: center; color: #a0a0a0;">${Math.max(0, row.boqTotal - row.actualTotal)}</td>
                <td class="col-concrete" style="text-align: center;">${row.concrete}</td>
                <td class="col-wooden" style="text-align: center;">${row.wooden}</td>
                <td class="col-progress" style="width: 70px;">
                    <div style="display: flex; align-items: center; gap: 4px;">
                        <div style="flex-grow: 1; height: 4px; background: #333; border-radius: 2px; overflow: hidden;">
                            <div style="width: ${Math.min(100, progress)}%; height: 100%; background: ${statusColor};"></div>
                        </div>
                        <span style="font-size: 0.8em; color: ${statusColor};">${progress.toFixed(0)}%</span>
                    </div>
                </td>
                <td class="col-status"><span style="font-size: 0.8em; padding: 1px 6px; border-radius: 8px; background: ${statusColor}20; color: ${statusColor}; border: 1px solid ${statusColor}40; white-space: nowrap;">${status}</span></td>
            `;
            tbody.appendChild(tr);
        });

        // 5. Update Info & Render Pagination Controls
        if (infoEl) infoEl.textContent = `Showing ${startIndex + 1}-${Math.min(endIndex, totalRows)} of ${totalRows} DTs`;
        renderPaginationControls(totalPages);
    }

    function renderPaginationControls(totalPages) {
        const container = document.getElementById('paginationControls');
        if (!container) return;

        container.innerHTML = '';
        if (totalPages <= 1) return;

        // Prev Button
        const prevBtn = document.createElement('button');
        prevBtn.className = 'page-btn';
        prevBtn.innerHTML = '&lt;'; // <
        prevBtn.disabled = currentPage === 1;
        prevBtn.onclick = () => {
            if (currentPage > 1) {
                currentPage--;
                renderDTTable();
            }
        };
        container.appendChild(prevBtn);

        // Page Numbers (Smart display: First, Last, Current +/- 1)
        const pagesToShow = new Set([1, totalPages, currentPage, currentPage - 1, currentPage + 1]);
        const sortedPages = [...pagesToShow].filter(p => p >= 1 && p <= totalPages).sort((a, b) => a - b);

        let lastPage = 0;
        sortedPages.forEach(p => {
            if (lastPage > 0 && p - lastPage > 1) {
                // Ellipsis
                const span = document.createElement('span');
                span.className = 'page-ellipsis';
                span.textContent = '...';
                span.style.color = '#64748b';
                container.appendChild(span);
            }

            const btn = document.createElement('button');
            btn.className = `page-btn ${p === currentPage ? 'active' : ''}`;
            btn.textContent = p;
            btn.onclick = () => {
                currentPage = p;
                renderDTTable();
            };
            container.appendChild(btn);
            lastPage = p;
        });

        // Next Button
        const nextBtn = document.createElement('button');
        nextBtn.className = 'page-btn';
        nextBtn.innerHTML = '&gt;'; // >
        nextBtn.disabled = currentPage === totalPages;
        nextBtn.onclick = () => {
            if (currentPage < totalPages) {
                currentPage++;
                renderDTTable();
            }
        };
        container.appendChild(nextBtn);
    }



    function renderFeederChart() {
        const counts = {};
        filteredData.forEach(d => {
            const val = d.Feeder || "Unknown";
            counts[val] = (counts[val] || 0) + 1;
        });

        // Top 10 Feeders
        const sorted = Object.entries(counts).sort((a, b) => a[1] - b[1]).slice(-10);
        const y = sorted.map(d => d[0]);
        const x = sorted.map(d => d[1]);

        const trace = {
            x: x,
            y: y,
            type: 'bar',
            orientation: 'h',
            marker: { color: '#8b5cf6' } // Purple
        };

        const layout = {
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            font: { color: '#e4e5e7' },
            margin: { l: 200, t: 30, b: 50, r: 20 },
            xaxis: { title: 'Count' },
            yaxis: { automargin: true }
        };

        const config = { responsive: true, displayModeBar: false };
        Plotly.newPlot('feederChart', [trace], layout, config);
    }

    // 7. Render Map (Leaflet)
    function renderMap() {
        if (!map) {
            // Init map if not exists
            // Center on Shomolu approx
            map = L.map('map').setView([6.536, 3.357], 15);
            // Add OSM tile layer
            L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                attribution: '&copy; OpenStreetMap contributors'
            }).addTo(map);
            markersLayer = L.layerGroup().addTo(map);
        }

        markersLayer.clearLayers();

        let count = 0;
        const limit = 3000; // Performance limit

        filteredData.forEach(d => {
            if (count > limit) return;

            const lat = parseFloat(d.Latitude);
            const lon = parseFloat(d.Longitude);

            if (!isNaN(lat) && !isNaN(lon)) {
                let color = '#a0a0a0';
                if (d.Vendor_Name === 'ETC Workforce') color = '#0EA5E9';
                if (d.Vendor_Name === 'Jesom Technology') color = '#EF4444';

                const marker = L.circleMarker([lat, lon], {
                    radius: 6,
                    fillColor: color,
                    color: '#fff',
                    weight: 1,
                    opacity: 1,
                    fillOpacity: 0.8
                });

                const popupContent = `
                    <div style="font-size: 0.9em; color: #333;">
                        <b>Pole ID:</b> ${d["Lt PoleSLRN"] || d["LT Pole No"] || "N/A"}<br>
                        <b>DT Name:</b> ${d["DT Name"] || "N/A"}<br>
                        <b>Vendor:</b> ${d.Vendor_Name}<br>
                        <b>Officer:</b> ${userFullNames[d.User] || d.User}<br>
                        <b>Coords:</b> ${lat.toFixed(5)}, ${lon.toFixed(5)}
                    </div>
                `;
                marker.bindPopup(popupContent);
                markersLayer.addLayer(marker);
                count++;
            }
        });

        // Fit Bounds if we have markers
        if (count > 0 && markersLayer.getLayers().length > 0) {
            const group = L.featureGroup(markersLayer.getLayers());
            // Add a slight delay or just fit
            try {
                map.fitBounds(group.getBounds());
            } catch (e) { console.warn("Bounds error", e); }
        }
    }

    function updateKeyInsights() {
        const container = document.getElementById('keyInsightsContent');
        if (!container) return;

        // Calculate Metrics
        // 1. Top Vendor
        const vendorCounts = {};
        filteredData.forEach(d => {
            const val = d["Vendor_Name"] || "Unassigned";
            vendorCounts[val] = (vendorCounts[val] || 0) + 1;
        });
        const topVendor = Object.entries(vendorCounts).sort((a, b) => b[1] - a[1])[0] || ['N/A', 0];

        // 2. Field Officer Stats (Top & Bottom)
        const userCounts = {};
        filteredData.forEach(d => {
            const val = d.User || "Unknown";
            userCounts[val] = (userCounts[val] || 0) + 1;
        });
        const sortedUsers = Object.entries(userCounts).sort((a, b) => b[1] - a[1]);

        const topUserEntry = sortedUsers[0] || ['N/A', 0];
        const topUser = userFullNames[topUserEntry[0]] || topUserEntry[0];

        const bottomUserEntry = sortedUsers[sortedUsers.length - 1] || ['N/A', 0];
        const bottomUser = userFullNames[bottomUserEntry[0]] || bottomUserEntry[0];
        // Find vendor for bottom user
        const bottomUserRecord = filteredData.find(d => d.User === bottomUserEntry[0]);
        const bottomVendor = bottomUserRecord ? (bottomUserRecord["Vendor_Name"] || 'Unknown') : 'Unknown';
        const bottomPct = filteredData.length > 0 ? ((bottomUserEntry[1] / filteredData.length) * 100).toFixed(1) : 0;

        // 3. Top Undertaking
        const utCounts = {};
        filteredData.forEach(d => {
            const val = d.Undertaking || "Unknown";
            utCounts[val] = (utCounts[val] || 0) + 1;
        });
        const sortedUt = Object.entries(utCounts).sort((a, b) => b[1] - a[1]);
        const topUt = sortedUt[0] || ['N/A', 0];
        const bottomUt = sortedUt[sortedUt.length - 1] || ['N/A', 0];

        // 4. Data Coverage (Date)
        const dates = filteredData.map(d => new Date(d["Date/timestamp"])).filter(d => !isNaN(d));
        let dateRange = "N/A";
        if (dates.length > 0) {
            const minDate = new Date(Math.min(...dates)).toLocaleDateString();
            const maxDate = new Date(Math.max(...dates)).toLocaleDateString();
            dateRange = `${minDate} - ${maxDate}`;
        }



        // Check if a specific user is selected
        const selectedUser = document.getElementById('userFilter').value;
        const showOfficerStats = selectedUser === 'All';

        let officerStatsHTML = '';
        if (showOfficerStats) {
            officerStatsHTML = `
            <div class="insight-item">
                <span class="insight-label">Top Field Officer</span>
                <span class="insight-value">${topUser}</span>
            </div>
            <div class="insight-item">
                <span class="insight-label">Bottom Field Officer</span>
                <div style="text-align: right;">
                    <div class="insight-value bottom-perf">${bottomUser}</div>
                    <div style="font-size: 0.8rem; color: var(--text-secondary);">${bottomVendor} (${bottomPct}%)</div>
                </div>
            </div>`;
        }

        container.innerHTML = `
            <div class="insight-item">
                <span class="insight-label">Running Vendor</span>
                <span class="insight-value highlight">${topVendor[0]} <span style="font-size:0.8em; color:var(--text-secondary);">(${((topVendor[1] / filteredData.length) * 100).toFixed(0)}%)</span></span>
            </div>
            ${officerStatsHTML}
            <div class="insight-item">
                <span class="insight-label">Highest Undertaking</span>
                <span class="insight-value">${topUt[0]}</span>
            </div>
            <div class="insight-item">
                <span class="insight-label">Lowest Undertaking</span>
                <span class="insight-value">${bottomUt[0]}</span>
            </div>
            <div class="insight-item">
                <span class="insight-label">Data Coverage</span>
                <span class="insight-value" style="font-size: 0.9rem;">${dateRange}</span>
            </div>
        `;
    }

    // Navigation Logic
    const navHome = document.getElementById('nav-home');
    const navDashboard = document.getElementById('nav-dashboard');
    const viewHome = document.getElementById('view-home');
    const viewDashboard = document.getElementById('view-dashboard');
    const dashboardSubLinks = document.getElementById('dashboard-sub-links');

    if (navHome && navDashboard) {
        navHome.addEventListener('click', (e) => {
            e.preventDefault();
            viewHome.classList.remove('hidden');
            viewDashboard.classList.add('hidden');
            navHome.classList.add('active');
            navDashboard.classList.remove('active');
            if (dashboardSubLinks) dashboardSubLinks.classList.add('hidden');
        });

        navDashboard.addEventListener('click', (e) => {
            e.preventDefault();
            viewHome.classList.add('hidden');
            viewDashboard.classList.remove('hidden');
            navHome.classList.remove('active');
            navDashboard.classList.add('active');
            if (dashboardSubLinks) dashboardSubLinks.classList.remove('hidden');

            // Trigger chart resize in case they were hidden
            window.dispatchEvent(new Event('resize'));
        });
    }


    // --- VARIANCE LOGIC & HELPERS ---



    function handleViewModeToggle(e) {
        viewMode = e.target.checked ? 'boq' : 'field';
        updateDashboard();
    }

    // Merge Function
    function calculateVariance() {
        // 1. Group Field Data by Feeder + DT
        // Key: "Feeder|DT Name"
        const fieldGroups = {};

        filteredData.forEach(d => {
            const feeder = (d.Feeder || "").trim().toUpperCase();
            const dt = (d["DT Name"] || "").trim().toUpperCase();
            const key = `${feeder}|${dt}`;

            if (!fieldGroups[key]) {
                fieldGroups[key] = {
                    feeder: d.Feeder,
                    dtName: d["DT Name"],
                    vendor: d.Vendor_Name,
                    actualTotal: 0,
                    actualGood: 0,
                    actualBad: 0,
                    users: new Set()
                };
            }
            fieldGroups[key].actualTotal++;
            if (d.Issue_Type === 'Good Condition') fieldGroups[key].actualGood++;
            else fieldGroups[key].actualBad++;
            fieldGroups[key].users.add(d.User);
        });

        // 2. Iterate BOQ and Merge
        // Apply Filters to BOQ Data as well (Feeder and DT only)
        const selectedFeeder = document.getElementById('feederFilter').value;
        const selectedDT = document.getElementById('dtFilter').value;

        const filteredBOQ = boqData.filter(boq => {
            if (selectedFeeder && selectedFeeder !== 'All' && boq["FEEDER NAME"] !== selectedFeeder) return false;
            if (selectedDT && selectedDT !== 'All' && boq["DT NAME"] !== selectedDT) return false;
            return true;
        });

        const merged = filteredBOQ.map(boq => {
            const feeder = (boq["FEEDER NAME"] || "").trim().toUpperCase();
            const dt = (boq["DT NAME"] || "").trim().toUpperCase();
            const key = `${feeder}|${dt}`;

            const field = fieldGroups[key] || { actualTotal: 0, actualGood: 0, actualBad: 0, users: new Set(), vendor: 'N/A' };

            const boqTotal = parseInt(boq["POLES Grand Total"]) || 0;
            const boqGood = parseInt(boq["GOOD"]) || 0;
            const boqBad = parseInt(boq["BAD"]) || 0;

            const variance = boqTotal > 0 ? ((field.actualTotal - boqTotal) / boqTotal * 100) : 0; // % Diff? Or just use raw diff?
            // User requested: "Variance (%)"
            // Formula: (Actual - BOQ) / BOQ * 100 usually. 
            // If Actual < BOQ, negative %. If Actual > BOQ, positive %.

            return {
                feeder: boq["FEEDER NAME"],
                dtName: boq["DT NAME"],
                vendor: field.vendor === 'N/A' ? 'Not Started' : field.vendor,
                boqTotal: boqTotal,
                actualTotal: field.actualTotal,
                boqGood,
                actualGood: field.actualGood,
                boqBad,
                actualBad: field.actualBad,
                variance: variance,
                users: Array.from(field.users)
            };
        });

        // Also include Field items that were NOT in BOQ? (New discoveries)
        // User didn't strictly ask, but good practice.
        // For simplicity, sticking to BOQ base as "Baseline BOQ" implies.

        return merged;
    }

    function renderVarianceCharts() {
        const mergedData = calculateVariance();

        // Chart 1: Target vs Actual (Bulleted Progres) - Top 10 Feeders or Global? 
        // User: "Feeders & DTs". Let's do Top 10 Feeders by Volume
        const feederStats = {};
        mergedData.forEach(d => {
            const f = d.feeder || "Unknown";
            if (!feederStats[f]) feederStats[f] = { boq: 0, act: 0 };
            feederStats[f].boq += d.boqTotal;
            feederStats[f].act += d.actualTotal;
        });

        const sortedFeeders = Object.entries(feederStats)
            .sort((a, b) => b[1].boq - a[1].boq)
            .slice(0, 10);

        const feederLabels = sortedFeeders.map(x => x[0]);
        const feederBoq = sortedFeeders.map(x => x[1].boq);
        const feederAct = sortedFeeders.map(x => x[1].act);

        // ApexChart Options for Target vs Actual
        const options1 = {
            series: [
                { name: 'Actual Captured', data: feederAct },
                { name: 'Total Target', data: feederBoq }
            ],
            chart: { type: 'bar', height: 400, toolbar: { show: false }, background: 'transparent' },
            plotOptions: {
                bar: {
                    horizontal: true,
                    dataLabels: { position: 'top' },
                }
            },
            colors: ['#10b981', 'rgba(16, 185, 129, 0.3)'], // Solid Green, Transparent Green
            dataLabels: {
                enabled: true,
                offsetX: -6,
                style: { fontSize: '12px', colors: ['#fff'] }
            },
            stroke: { show: true, width: 1, colors: ['#fff'] },
            xaxis: { title: { text: 'Number of Poles', style: { color: '#a0a0a0' } }, labels: { style: { colors: '#a0a0a0' } } },
            yaxis: { labels: { style: { colors: '#fff' } } },
            theme: { mode: 'dark' },
            grid: { borderColor: '#373a40' }
        };

        const chart1El = document.querySelector("#targetActualChart");
        chart1El.innerHTML = ""; // Clear
        const chart1 = new ApexCharts(chart1El, options1);
        chart1.render();

        // Chart 2: Pole Health Reconciliation (Grouped Bar) - Top 10 DTs
        const topDTs = mergedData
            .sort((a, b) => b.boqTotal - a.boqTotal)
            .slice(0, 10);

        const dtLabels = topDTs.map(d => d.dtName);

        const options2 = {
            series: [
                { name: 'Total Bad', data: topDTs.map(d => d.boqBad) },
                { name: 'Actual Bad', data: topDTs.map(d => d.actualBad) },
                { name: 'Total Good', data: topDTs.map(d => d.boqGood) },
                { name: 'Actual Good', data: topDTs.map(d => d.actualGood) }
            ],
            chart: { type: 'bar', height: 400, toolbar: { show: false }, background: 'transparent' },
            colors: ['rgba(239, 68, 68, 0.4)', '#ef4444', 'rgba(16, 185, 129, 0.4)', '#10b981'], // Transp Red, Solid Red, Transp Green, Solid Green
            plotOptions: {
                bar: { horizontal: false, columnWidth: '55%', endingShape: 'rounded' }
            },
            dataLabels: { enabled: false },
            xaxis: { categories: dtLabels, labels: { style: { colors: '#a0a0a0' } } },
            yaxis: { title: { text: 'Count', style: { color: '#a0a0a0' } }, labels: { style: { colors: '#a0a0a0' } } },
            theme: { mode: 'dark' },
            grid: { borderColor: '#373a40' },
            legend: { labels: { colors: '#fff' } }
        };

        const chart2El = document.querySelector("#poleHealthChart");
        chart2El.innerHTML = "";
        const chart2 = new ApexCharts(chart2El, options2);
        chart2.render();
    }


    // 8. Render Strategic Recommendations (Dynamic)
    function renderStrategicRecommendations() {
        // Use globalData to ensure recommendations stand regardless of transient filters
        // unless we want them to reflect the filtered view. Strategic usually implies overall.
        // Let's use globalData for stability.

        const vendors = ['ETC Workforce', 'Jesom Technology'];

        vendors.forEach(vendor => {
            // Filter Data for Vendor
            const vData = globalData.filter(d => d.Vendor_Name === vendor);

            if (vData.length === 0) return; // No data, skip

            // --- Metrics Calculation ---

            // 1. Run Rate
            // distinct dates
            const dates = new Set(vData.map(d => d["Date/timestamp"] ? d["Date/timestamp"].split(' ')[0] : ''));
            const activeDays = dates.size || 1;
            const totalRecords = vData.length;
            const avgRate = (totalRecords / activeDays).toFixed(0); // Integer for readability

            // 2. Coverage (Undertakings)
            const activeUTs = new Set(vData.map(d => d.Undertaking)).size;

            // 3. Quality (Defect Rate)
            const badPoles = vData.filter(d => d.Issue_Type && d.Issue_Type !== 'Good Condition').length;
            const defectPct = ((badPoles / totalRecords) * 100).toFixed(1);

            // 4. Synchronization (Last Date)
            // Assumes yyyy-mm-dd or similar sortable via Date parse
            const sortedDates = Array.from(dates).sort();
            const lastDateISO = sortedDates[sortedDates.length - 1];
            const lastDateObj = lastDateISO ? new Date(lastDateISO) : new Date();
            const today = new Date();
            // Diff in days
            const diffTime = Math.abs(today - lastDateObj);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));


            // --- Logic & Text Formatting ---

            let status = 'On Track';
            let statusClass = 'status-good';

            // Thresholds
            const TARGET_RATE = 50;

            if (avgRate < TARGET_RATE) {
                status = 'Requires Attention';
                statusClass = 'status-attention';
            }

            // Rec 1: Velocity
            let rec1 = {};
            if (avgRate < 30) {
                rec1.icon = '🚀';
                rec1.title = 'Accelerate Deployment';
                rec1.text = `Current velocity (${avgRate}/day) is critically low. Immediate resource scale-up is required to meet the daily target of ${TARGET_RATE}.`;
            } else if (avgRate < TARGET_RATE) {
                rec1.icon = '⚠️';
                rec1.title = 'Increase Pace';
                rec1.text = `Current velocity (${avgRate}/day) is approaching target but still falls short. Consider extending operational hours.`;
            } else {
                rec1.icon = '⭐';
                rec1.title = 'Sustain Momentum';
                rec1.text = `Strong performance with a velocity of ${avgRate}/day. Keep this consistency to ensure project timelines are met.`;
            }

            // Rec 2: Coverage
            let rec2 = {};
            // Logic: Is it clustered?
            // Simple heuristic: If high volume but low UT count -> Clustered.
            if (activeUTs < 2 && totalRecords > 100) {
                rec2.icon = '📍';
                rec2.title = 'Expand Coverage';
                rec2.text = `High activity concentration detected in only ${activeUTs} Undertaking. Redeploy teams to under-served areas to avoid data gaps.`;
            } else {
                rec2.icon = '🗺️';
                rec2.title = 'Balanced Coverage';
                rec2.text = `Active across ${activeUTs} Undertakings. Continue maintaining balanced visibility across the network.`;
            }

            // Rec 3: Quality / Data Sync
            let rec3 = {};
            // Prio 1: specific defect issues
            if (defectPct > 20) {
                rec3.icon = '📝';
                rec3.title = 'Enhanced Reporting';
                rec3.text = `High defect rate (${defectPct}%). Ensure engineering validation is performed to confirm the accuracy of 'Bad' pole tags.`;
            } else if (defectPct < 2) {
                rec3.icon = '🎯';
                rec3.title = 'Precision Check';
                rec3.text = `Defect rate is unusually low (${defectPct}%). Conduct spot checks to ensure defects are not being overlooked.`;
            } else if (diffDays > 3) {
                rec3.icon = '🔄';
                rec3.title = 'Data Synchronization';
                rec3.text = `Data lag detected (last active: ${lastDateISO}). Enforce daily sync protocols to maintain real-time dashboard accuracy.`;
            } else {
                rec3.icon = '✅';
                rec3.title = 'Quality Assurance';
                rec3.text = `Data quality appears healthy (Defect Rate: ${defectPct}%). Continue standard verification procedures.`;
            }

            // --- Render ---
            const idKey = vendor.split(' ')[0].toLowerCase();
            const badge = document.getElementById(`status-badge-${idKey}`);
            const content = document.getElementById(`rec-content-${idKey}`);

            if (badge) {
                badge.textContent = status;
                badge.className = `status-badge ${statusClass}`;
            }

            if (content) {
                content.innerHTML = `
                    <div class="rec-item">
                        <h4>${rec1.icon} ${rec1.title}</h4>
                        <p>${rec1.text}</p>
                    </div>
                    <div class="rec-item">
                        <h4>${rec2.icon} ${rec2.title}</h4>
                        <p>${rec2.text}</p>
                    </div>
                    <div class="rec-item">
                        <h4>${rec3.icon} ${rec3.title}</h4>
                        <p>${rec3.text}</p>
                    </div>
                `;
            }
        });
    }

    // --- Event Listeners ---
    const resetFiltersBtn = document.getElementById('resetFiltersBtn');
    if (resetFiltersBtn) {
        resetFiltersBtn.addEventListener('click', resetFilters);
    }

    const downloadCsvBtn = document.getElementById('downloadCSV');
    if (downloadCsvBtn) {
        downloadCsvBtn.addEventListener('click', () => {
            const data = getEnhancedDTData();
            // Filter by current search input
            const searchVal = (document.getElementById('dtSearchInput')?.value || '').toLowerCase();
            const filtered = data.filter(item => {
                if (!searchVal) return true;
                return (
                    (item.dtName || '').toLowerCase().includes(searchVal) ||
                    (item.vendor || '').toLowerCase().includes(searchVal) ||
                    (item.feeder || '').toLowerCase().includes(searchVal) ||
                    (item.bu || '').toLowerCase().includes(searchVal) ||
                    (item.undertaking || '').toLowerCase().includes(searchVal) ||
                    item.users.some(u => String(userFullNames[u] || u || '').toLowerCase().includes(searchVal))
                );
            });

            if (!filtered || filtered.length === 0) {
                alert("No data to export");
                return;
            }

            const headers = ["DT Name", "Feeder", "BU", "Undertaking", "Vendor", "Users", "BOQ Total", "Actual Total", "Gap", "Concrete", "Wooden"];

            const csvRows = [headers.join(',')];

            filtered.forEach(row => {
                const values = [
                    `"${row.dtName || ''}"`,
                    `"${row.feeder || ''}"`,
                    `"${row.bu || ''}"`,
                    `"${row.undertaking || ''}"`,
                    `"${row.vendor || ''}"`,
                    `"${(row.users || []).map(u => userFullNames[u] || u).join('; ')}"`,
                    row.boqTotal || 0,
                    row.actualTotal || 0,
                    Math.max(0, (row.boqTotal || 0) - (row.actualTotal || 0)),
                    row.concrete || 0,
                    row.wooden || 0
                ];
                csvRows.push(values.join(','));
            });

            const csvContent = csvRows.join('\n');
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement("a");
            const url = URL.createObjectURL(blob);
            link.setAttribute("href", url);
            link.setAttribute("download", "dt_analysis_export.csv");
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        });
    }

});