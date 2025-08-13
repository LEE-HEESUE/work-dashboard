// 전역 변수들
let weeklyDataSets = {};
let availableWeeks = [];
let selectedWeek = '';
let currentWeekData = [];
let weekSearchTerm = '';

// GitHub 관련 설정
let isAdminMode = false;
let githubConfig = {
    username: '',
    repo: '',
    token: ''
};

// 공개 Repository 정보 (일반 사용자도 데이터를 볼 수 있도록)
const PUBLIC_GITHUB_CONFIG = {
    username: 'LEE-HEESUE', // 당신의 GitHub 사용자명
    repo: 'work-dashboard'
};
let organizations = {
    소속1: [], 소속2: [], 소속3: [], 트라이브: [], 소속4: []
};
let filters = {
    소속1: [], 소속2: [], 소속3: [], 트라이브: [], 소속4: []
};
let searchTerms = {
    소속1: '', 소속2: '', 소속3: '', 트라이브: '', 소속4: ''
};
let employeeFilters = {
    소속1: [], 소속2: [], 소속3: [], 트라이브: [], 소속4: [], 고용구분: []
};
let employeeSearch = '';
let sortConfig = { key: null, direction: 'asc' };
let employeeSortConfig = { column: null, direction: 'asc' };

// 알럿 관련 변수들
let alertData = [];
let filteredAlertData = [];
let alertSearchTerm = '';
let alertFilters = {
    weekly: true,
    coretime: true,
    daily: true
};

// 근무 미달 대상자 관련 변수들
let shortageData = {
    daily: [],
    weekly: [],
    coretime: []
};
let currentShortageTab = 'daily';
let shortageSortConfig = {
    daily: { column: null, direction: 'asc' },
    weekly: { column: null, direction: 'asc' },
    coretime: { column: null, direction: 'asc' }
};

// 초과근무 테이블 관련 변수들
let overtimeSortConfig = {
    regular: { column: null, direction: 'asc' },
    contract: { column: null, direction: 'asc' }
};

// 차트 인스턴스들
let workTimeChart = null;
let overtimeChart = null;

// DOM 요소들
const fileInput = document.getElementById('fileInput');
const weekSelectorContainer = document.getElementById('weekSelectorContainer');
const weekSearch = document.getElementById('weekSearch');
const weekOptions = document.getElementById('weekOptions');
const selectedWeekInfo = document.getElementById('selectedWeekInfo');
const selectedWeekDisplay = document.getElementById('selectedWeekDisplay');
const errorMessage = document.getElementById('errorMessage');
const errorText = document.getElementById('errorText');
const loadingSpinner = document.getElementById('loadingSpinner');
const dashboardContent = document.getElementById('dashboardContent');
const welcomeMessage = document.getElementById('welcomeMessage');

// 시간 변환 함수들
function convertToHours(days) {
    return (days || 0) * 24;
}

function formatHours(days) {
    if (!days || isNaN(days)) return '00시간 00분';
    const totalMinutes = Math.round(days * 24 * 60);
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    return `${hours.toString().padStart(2, '0')}시간 ${minutes.toString().padStart(2, '0')}분`;
}

// 에러 표시 함수
function showError(message) {
    errorText.textContent = message;
    errorMessage.classList.remove('hidden');
    setTimeout(() => {
        errorMessage.classList.add('hidden');
    }, 5000);
}

// 성공 메시지 표시
function showSuccessMessage(message) {
    // 기존 성공 메시지 제거
    const existingSuccess = document.getElementById('successMessage');
    if (existingSuccess) {
        existingSuccess.remove();
    }

    // 새 성공 메시지 생성
    const successDiv = document.createElement('div');
    successDiv.id = 'successMessage';
    successDiv.className = 'mb-6 bg-green-50 border border-green-200 rounded-lg p-4';
    successDiv.innerHTML = `
        <div class="flex items-center">
            <div class="text-green-600 mr-2">✅</div>
            <p class="text-green-700">${message}</p>
        </div>
    `;

    // 에러 메시지 뒤에 삽입
    errorMessage.parentNode.insertBefore(successDiv, errorMessage.nextSibling);

    // 5초 후 자동 제거
    setTimeout(() => {
        if (successDiv.parentNode) {
            successDiv.remove();
        }
    }, 5000);
}

// 로딩 상태 관리
function setLoading(isLoading) {
    if (isLoading) {
        loadingSpinner.classList.remove('hidden');
        dashboardContent.style.display = 'none';
        welcomeMessage.style.display = 'none';
    } else {
        loadingSpinner.classList.add('hidden');
    }
}

// LocalStorage 관련 함수들
const STORAGE_KEY = 'workTimeData';

function saveToStorage(weekName, data) {
    try {
        const existingData = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
        existingData[weekName] = {
            data: data,
            savedAt: new Date().toISOString(),
            dataCount: data.length
        };
        localStorage.setItem(STORAGE_KEY, JSON.stringify(existingData));
        return true;
    } catch (error) {
        console.error('저장 실패:', error);
        return false;
    }
}

function loadFromStorage() {
    try {
        const data = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
        return data;
    } catch (error) {
        console.error('로드 실패:', error);
        return {};
    }
}

function deleteWeekFromStorage(weekName) {
    try {
        const existingData = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
        delete existingData[weekName];
        localStorage.setItem(STORAGE_KEY, JSON.stringify(existingData));
        return true;
    } catch (error) {
        console.error('삭제 실패:', error);
        return false;
    }
}

// GitHub 설정 관리
const GITHUB_CONFIG_KEY = 'githubConfig';

function saveGithubConfig(username, repo, token) {
    const config = { username, repo, token };
    localStorage.setItem(GITHUB_CONFIG_KEY, JSON.stringify(config));
    githubConfig = config;
    return true;
}

function loadGithubConfig() {
    try {
        const config = JSON.parse(localStorage.getItem(GITHUB_CONFIG_KEY) || '{}');
        if (config.username && config.repo && config.token) {
            githubConfig = config;
            return true;
        }
        return false;
    } catch (error) {
        console.error('GitHub 설정 로드 실패:', error);
        return false;
    }
}

// 관리자 모드 전환
function toggleAdminMode(enable) {
    isAdminMode = enable;
    updateUIForMode();
}

function updateUIForMode() {
    const userModeIndicator = document.getElementById('userModeIndicator');
    const fileUploadSection = document.querySelector('.file-upload-area').parentElement;
    const deleteButtons = document.querySelectorAll('#deleteSelectedWeek');
    
    if (isAdminMode) {
        userModeIndicator.textContent = '관리자 모드';
        userModeIndicator.className = 'text-sm px-3 py-1 rounded-full bg-green-100 text-green-700';
        if (fileUploadSection) fileUploadSection.style.display = 'block';
        deleteButtons.forEach(btn => btn.style.display = 'inline-flex');
    } else {
        userModeIndicator.textContent = '읽기 전용 모드';
        userModeIndicator.className = 'text-sm px-3 py-1 rounded-full bg-gray-100 text-gray-600';
        if (fileUploadSection) fileUploadSection.style.display = 'none';
        deleteButtons.forEach(btn => btn.style.display = 'none');
    }
}

// GitHub API 함수들
async function githubApiRequest(endpoint, method = 'GET', data = null) {
    const url = `https://api.github.com/repos/${githubConfig.username}/${githubConfig.repo}/${endpoint}`;
    
    const options = {
        method,
        headers: {
            'Authorization': `token ${githubConfig.token}`,
            'Accept': 'application/vnd.github.v3+json',
            'Content-Type': 'application/json'
        }
    };
    
    if (data && (method === 'POST' || method === 'PUT')) {
        options.body = JSON.stringify(data);
    }
    
    const response = await fetch(url, options);
    
    if (!response.ok) {
        throw new Error(`GitHub API 오류: ${response.status} ${response.statusText}`);
    }
    
    return response.json();
}

// 엑셀 데이터를 JSON으로 변환
function convertExcelToJson(weeklyData, shortageData, weekName) {
    return {
        weekName,
        uploadDate: new Date().toISOString(),
        weeklyData,
        shortageData,
        metadata: {
            totalEmployees: weeklyData.length,
            version: '1.0'
        }
    };
}

// GitHub에 데이터 업로드
async function uploadToGithub(weekName, processedData) {
    try {
        const fileName = `data/${weekName.replace(/[^a-zA-Z0-9가-힣]/g, '_')}.json`;
        
        // 기존 파일이 있는지 확인
        let sha = null;
        try {
            const existingFile = await githubApiRequest(`contents/${fileName}`);
            sha = existingFile.sha;
        } catch (error) {
            // 파일이 없으면 새로 생성
        }
        
        const content = btoa(unescape(encodeURIComponent(JSON.stringify(processedData, null, 2))));
        
        const requestData = {
            message: `Upload ${weekName} data`,
            content,
            ...(sha && { sha })
        };
        
        await githubApiRequest(`contents/${fileName}`, 'PUT', requestData);
        
        // 주차 목록 업데이트
        await updateWeekListFile();
        
        return true;
    } catch (error) {
        console.error('GitHub 업로드 실패:', error);
        throw error;
    }
}

// 주차 목록 파일 업데이트
async function updateWeekListFile() {
    try {
        // data 폴더의 모든 JSON 파일 목록 가져오기
        const contents = await githubApiRequest('contents/data');
        const weeks = contents
            .filter(file => file.name.endsWith('.json') && file.name !== 'week_list.json')
            .map(file => file.name.replace('.json', '').replace(/_/g, ' '));
        
        const weekListData = {
            weeks: weeks.sort().reverse(),
            lastUpdated: new Date().toISOString()
        };
        
        // week_list.json 파일 업데이트
        let sha = null;
        try {
            const existingFile = await githubApiRequest('contents/data/week_list.json');
            sha = existingFile.sha;
        } catch (error) {
            // 파일이 없으면 새로 생성
        }
        
        const content = btoa(unescape(encodeURIComponent(JSON.stringify(weekListData, null, 2))));
        
        const requestData = {
            message: 'Update week list',
            content,
            ...(sha && { sha })
        };
        
        await githubApiRequest('contents/data/week_list.json', 'PUT', requestData);
        
    } catch (error) {
        console.error('주차 목록 업데이트 실패:', error);
        throw error;
    }
}

// GitHub에서 데이터 로드 (일반 사용자도 접근 가능)
async function loadFromGithubPublic(username, repo) {
    try {
        // 주차 목록 가져오기
        const weekListResponse = await fetch(`https://raw.githubusercontent.com/${username}/${repo}/main/data/week_list.json`);
        
        if (!weekListResponse.ok) {
            console.log('GitHub에서 데이터를 찾을 수 없습니다.');
            return {};
        }
        
        const weekListData = await weekListResponse.json();
        const weeks = weekListData.weeks || [];
        
        const githubData = {};
        
        // 각 주차 데이터 로드
        for (const week of weeks) {
            try {
                const fileName = week.replace(/[^a-zA-Z0-9가-힣]/g, '_');
                const dataResponse = await fetch(`https://raw.githubusercontent.com/${username}/${repo}/main/data/${fileName}.json`);
                
                if (dataResponse.ok) {
                    const weekData = await dataResponse.json();
                    githubData[week] = {
                        data: weekData.weeklyData,
                        savedAt: weekData.uploadDate,
                        dataCount: weekData.weeklyData.length,
                        shortageData: weekData.shortageData
                    };
                }
            } catch (error) {
                console.error(`${week} 데이터 로드 실패:`, error);
            }
        }
        
        return githubData;
        
    } catch (error) {
        console.error('GitHub 데이터 로드 실패:', error);
        return {};
    }
}

// GitHub에서 데이터 로드 (관리자용 - 인증 필요)
async function loadFromGithub() {
    return loadFromGithubPublic(githubConfig.username, githubConfig.repo);
}

// 파일 업로드 핸들러
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
        setLoading(true);
        errorMessage.classList.add('hidden');

        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);

        const weeklySheet = workbook.Sheets['2. 근무기록(주간)'];
        if (!weeklySheet) {
            throw new Error('\'2. 근무기록(주간)\' 시트를 찾을 수 없습니다.');
        }

        const weeklyData = XLSX.utils.sheet_to_json(weeklySheet);

        // 알럿 관련 시트들도 함께 로드
        console.log('사용 가능한 시트들:', Object.keys(workbook.Sheets));
        console.log('각 시트의 첫 번째 행 데이터 확인:');
        
        // 각 시트의 컬럼 구조 확인
        ['수정요청리스트(일근무)', '수정요청리스트(주근무)', '수정요청리스트(코어타임)'].forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            if (sheet) {
                console.log(`\n=== ${sheetName} 시트 분석 ===`);
                const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                console.log(`${sheetName} 시트 헤더:`, data[0]);
                
                // 전체 데이터 확인
                const jsonData = XLSX.utils.sheet_to_json(sheet);
                console.log(`${sheetName} 전체 행 수:`, jsonData.length);
                
                if (jsonData.length > 0) {
                    console.log(`${sheetName} 첫 번째 데이터:`, jsonData[0]);
                    
                    // 알럿 여부 컬럼 확인
                    const alertColumns = ['알럿 여부', '알럿', '검토대상', '알림'];
                    alertColumns.forEach(col => {
                        if (jsonData[0].hasOwnProperty(col)) {
                            console.log(`${sheetName}에서 '${col}' 컬럼 발견!`);
                            console.log(`'${col}' 컬럼의 모든 값들:`, jsonData.map(row => row[col]));
                        }
                    });
                }
            } else {
                console.log(`❌ ${sheetName} 시트를 찾을 수 없습니다`);
            }
        });
        
        loadAlertSheets(workbook);
        loadShortageData(workbook);

        // 파일명을 그대로 주차 이름으로 사용 (확장자만 제거)
        const weekName = file.name.replace(/\.[^/.]+$/, "");

        // 주차별 데이터에 추가
        weeklyDataSets[weekName] = weeklyData;

        // LocalStorage에 저장
        if (saveToStorage(weekName, weeklyData)) {
            let successMessage = `${weekName} 데이터가 로컬에 저장되었습니다! (${weeklyData.length}개 행)`;
            
            // 관리자 모드이고 GitHub 설정이 있으면 GitHub에도 업로드
            if (isAdminMode && githubConfig.username && githubConfig.repo && githubConfig.token) {
                try {
                    const processedData = convertExcelToJson(weeklyData, shortageData, weekName);
                    await uploadToGithub(weekName, processedData);
                    successMessage = `${weekName} 데이터가 GitHub에 업로드되었습니다! (${weeklyData.length}개 행)`;
                } catch (error) {
                    showError(`GitHub 업로드 실패: ${error.message}. 로컬에만 저장되었습니다.`);
                }
            }
            
            showSuccessMessage(successMessage);
        }

        // 사용 가능한 주차 목록 업데이트
        availableWeeks = Object.keys(weeklyDataSets).sort().reverse();
        updateWeekOptions();

        // 새로 업로드한 주차를 선택
        handleWeekSelection(weekName);

        // 파일 입력 초기화
        fileInput.value = '';

    } catch (error) {
        showError(`파일 로드 실패: ${error.message}`);
    } finally {
        setLoading(false);
    }
}

// 주차 선택 옵션 업데이트 
function updateWeekOptions() {
    updateWeekList();
    
    if (availableWeeks.length > 0) {
        weekSelectorContainer.style.display = 'block';
        document.getElementById('weekCount').textContent = `${availableWeeks.length}개의 주차 데이터가 있습니다.`;
    }
}

// 주차 목록 업데이트 (검색 필터 적용)
function updateWeekList() {
    const searchTerm = weekSearchTerm.toLowerCase();
    const filteredWeeks = availableWeeks.filter(week =>
        week.toLowerCase().includes(searchTerm)
    );

    weekOptions.innerHTML = '';
    
    if (filteredWeeks.length > 0) {
        filteredWeeks.forEach(week => {
            const weekItem = document.createElement('div');
            weekItem.className = `week-option cursor-pointer p-2 rounded hover:bg-blue-50 text-sm ${
                selectedWeek === week ? 'bg-blue-100 text-blue-800 font-medium' : 'text-gray-700'
            }`;
            weekItem.innerHTML = `
                <div class="flex items-center justify-between">
                    <span>${week}</span>
                    ${selectedWeek === week ? '<i data-lucide="check" class="text-blue-600" size="16"></i>' : ''}
                </div>
            `;
            
            weekItem.addEventListener('click', () => {
                handleWeekSelection(week);
            });
            
            weekOptions.appendChild(weekItem);
        });
    } else {
        weekOptions.innerHTML = '<div class="text-xs text-gray-500 p-2">검색 결과 없음</div>';
    }
    
    // Lucide 아이콘 다시 초기화
    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }
}

// 주차 선택 처리
function handleWeekSelection(weekName) {
    selectedWeek = weekName;
    selectedWeekDisplay.textContent = weekName;
    selectedWeekInfo.style.display = 'block';
    
    // 주차 목록 다시 렌더링 (선택 상태 업데이트)
    updateWeekList();
    
    // 데이터 로드
    const weekData = weeklyDataSets[weekName] || [];
    loadWeekData(weekData);
}

// 주차 변경 핸들러 (호환성 유지)
function handleWeekChange(weekName) {
    handleWeekSelection(weekName);
}

// 선택된 주차 삭제
function deleteSelectedWeek() {
    if (!selectedWeek) {
        showError('삭제할 주차가 선택되지 않았습니다.');
        return;
    }

    // 확인 대화상자
    if (!confirm(`"${selectedWeek}" 주차 데이터를 정말 삭제하시겠습니까?\n\n이 작업은 되돌릴 수 없습니다.`)) {
        return;
    }

    try {
        // localStorage에서 삭제
        if (deleteWeekFromStorage(selectedWeek)) {
            // 메모리에서도 삭제
            delete weeklyDataSets[selectedWeek];
            
            // 사용 가능한 주차 목록 업데이트
            availableWeeks = Object.keys(weeklyDataSets).sort().reverse();
            
            // 성공 메시지
            showSuccessMessage(`"${selectedWeek}" 주차 데이터가 삭제되었습니다.`);
            
            // UI 업데이트
            if (availableWeeks.length > 0) {
                // 다른 주차가 있으면 가장 최근 주차 선택
                const latestWeek = availableWeeks[0];
                handleWeekSelection(latestWeek);
            } else {
                // 모든 주차가 삭제되었으면 초기 상태로
                selectedWeek = '';
                currentWeekData = [];
                selectedWeekInfo.style.display = 'none';
                weekSelectorContainer.style.display = 'none';
                dashboardContent.style.display = 'none';
                welcomeMessage.style.display = 'block';
            }
            
            updateWeekOptions();
            
        } else {
            showError('주차 데이터 삭제에 실패했습니다.');
        }
    } catch (error) {
        showError(`삭제 중 오류가 발생했습니다: ${error.message}`);
    }
}

// 저장된 데이터 초기 로드
async function initializeApp() {
    // GitHub 설정 로드
    const hasGithubConfig = loadGithubConfig();
    
    if (hasGithubConfig) {
        toggleAdminMode(true);
    } else {
        toggleAdminMode(false);
    }
    
    let storedData = {};
    
    try {
        setLoading(true);
        
        // 일반 사용자든 관리자든 공개 GitHub에서 데이터 로드 시도
        if (PUBLIC_GITHUB_CONFIG.username) {
            console.log('공개 GitHub에서 데이터 로드 시도...');
            storedData = await loadFromGithubPublic(PUBLIC_GITHUB_CONFIG.username, PUBLIC_GITHUB_CONFIG.repo);
            
            if (Object.keys(storedData).length > 0) {
                console.log('GitHub에서 데이터를 로드했습니다.');
            } else {
                console.log('GitHub에 데이터가 없습니다. 로컬 데이터를 확인합니다.');
                storedData = loadFromStorage();
            }
        } else {
            console.log('GitHub 설정이 없습니다. 로컬 데이터를 사용합니다.');
            storedData = loadFromStorage();
        }
        
    } catch (error) {
        console.log('데이터 로드 실패, 로컬 데이터를 사용합니다:', error.message);
        storedData = loadFromStorage();
    } finally {
        setLoading(false);
    }
    
    // 메모리에 로드
    Object.keys(storedData).forEach(weekName => {
        weeklyDataSets[weekName] = storedData[weekName].data;
        // GitHub에서 온 데이터는 shortageData도 함께 저장
        if (storedData[weekName].shortageData) {
            // shortageData 처리 로직 필요시 여기 추가
        }
    });
    
    availableWeeks = Object.keys(storedData).sort().reverse();
    updateWeekOptions();
    
    // 가장 최근 주차 자동 선택
    if (availableWeeks.length > 0) {
        const latestWeek = availableWeeks[0];
        handleWeekSelection(latestWeek);
    }
}

// 주차 데이터 로드
function loadWeekData(weekData) {
    currentWeekData = weekData;

    // 조직 목록 생성
    ['소속1', '소속2', '소속3', '트라이브', '소속4'].forEach(key => {
        organizations[key] = [...new Set(weekData.map(row => row[key]).filter(x => x))].sort();
    });

    // UI 업데이트
    updateFilters();
    updateStatistics();
    updateCharts();
    updateEmployeeTable();
    updateOvertimeTables(); // 초과근무 테이블 업데이트 추가
    updateAlerts();
    updateShortageSection(); // 알럿 업데이트 추가

    // 대시보드 표시
    welcomeMessage.style.display = 'none';
    dashboardContent.style.display = 'block';
}

// 필터 UI 업데이트
function updateFilters() {
    const container = document.getElementById('filtersContainer');
    container.innerHTML = '';

    Object.entries(organizations).forEach(([key, options]) => {
        const filterDiv = document.createElement('div');
        filterDiv.innerHTML = `
            <label class="block text-xs font-medium text-gray-700 mb-1">${key}</label>
            <input 
                type="text" 
                placeholder="${key} 검색..." 
                class="filter-search w-full p-1 mb-1 border border-gray-300 rounded text-xs focus:ring-1 focus:ring-blue-500"
                data-key="${key}"
            />
            <div class="filter-options border border-gray-300 rounded p-1 max-h-20 overflow-y-auto bg-white" data-key="${key}">
            </div>
            <div class="filter-count text-xs text-blue-600 mt-1" data-key="${key}" style="display: none;">
                0개 선택됨
            </div>
        `;
        container.appendChild(filterDiv);

        updateFilterOptions(key);
    });

    // 검색 이벤트 리스너
    document.querySelectorAll('.filter-search').forEach(input => {
        input.addEventListener('input', (e) => {
            const key = e.target.dataset.key;
            searchTerms[key] = e.target.value;
            updateFilterOptions(key);
        });
    });

    // 근무 미달 탭 이벤트 리스너들
    document.querySelectorAll('.shortage-tab').forEach(tab => {
        tab.addEventListener('click', (e) => {
            const tabKey = e.currentTarget.dataset.tab;
            switchShortageTab(tabKey);
        });
    });

    // 필터 초기화 버튼
    document.getElementById('clearFilters').addEventListener('click', clearFilters);
}

// 필터 옵션 업데이트
function updateFilterOptions(key) {
    const container = document.querySelector(`.filter-options[data-key="${key}"]`);
    const searchTerm = (searchTerms[key] || '').toLowerCase();
    const filteredOptions = organizations[key].filter(option =>
        option.toLowerCase().includes(searchTerm)
    );

    container.innerHTML = '';
    
    if (filteredOptions.length > 0) {
        filteredOptions.forEach(option => {
            const label = document.createElement('label');
            label.className = 'filter-checkbox flex items-center space-x-1 hover:bg-gray-50 p-0.5 rounded cursor-pointer';
            label.innerHTML = `
                <input 
                    type="checkbox" 
                    class="rounded border-gray-300 text-blue-600 w-3 h-3"
                    data-key="${key}"
                    data-value="${option}"
                    ${filters[key].includes(option) ? 'checked' : ''}
                />
                <span class="text-xs text-gray-700 truncate">${option}</span>
            `;
            container.appendChild(label);

            // 체크박스 이벤트
            label.querySelector('input').addEventListener('change', (e) => {
                handleFilterChange(key, option, e.target.checked);
            });
        });
    } else {
        container.innerHTML = '<div class="text-xs text-gray-500 p-1">검색 결과 없음</div>';
    }

    // 선택된 항목 수 업데이트
    const countElement = document.querySelector(`.filter-count[data-key="${key}"]`);
    if (filters[key].length > 0) {
        countElement.textContent = `${filters[key].length}개 선택됨`;
        countElement.style.display = 'block';
    } else {
        countElement.style.display = 'none';
    }
}

// 필터 변경 핸들러
function handleFilterChange(key, value, isChecked) {
    if (isChecked) {
        if (!filters[key].includes(value)) {
            filters[key].push(value);
        }
    } else {
        filters[key] = filters[key].filter(v => v !== value);
    }

    updateFilterOptions(key);
    updateStatistics();
    updateCharts();
    updateEmployeeTable();
    updateOvertimeTables(); // 초과근무 테이블도 업데이트
    updateShortageSection(); // 근무 미달 섹션도 업데이트
}

// 필터 초기화
function clearFilters() {
    filters = { 소속1: [], 소속2: [], 소속3: [], 트라이브: [], 소속4: [] };
    searchTerms = { 소속1: '', 소속2: '', 소속3: '', 트라이브: '', 소속4: '' };
    
    // 검색 입력 초기화
    document.querySelectorAll('.filter-search').forEach(input => {
        input.value = '';
    });

    // 필터 옵션 업데이트
    Object.keys(organizations).forEach(key => {
        updateFilterOptions(key);
    });

    updateStatistics();
    updateCharts();
    updateEmployeeTable();
    updateOvertimeTables(); // 초과근무 테이블도 업데이트
}

// 필터링된 데이터 가져오기
function getFilteredData() {
    return currentWeekData.filter(row => {
        return Object.entries(filters).every(([key, selectedValues]) => {
            if (!selectedValues.length) return true;
            const rowValue = row[key];
            return rowValue && selectedValues.some(value =>
                rowValue.toString().toLowerCase().includes(value.toLowerCase())
            );
        });
    });
}

// 통계 업데이트
function updateStatistics() {
    const filteredData = getFilteredData();
    const nonExecutives = filteredData.filter(row => row['고용구분'] !== '임원');
    const regularEmployees = nonExecutives.filter(row => row['고용구분'] === '정규');
    const fullTimeContractors = nonExecutives.filter(row =>
        row['고용구분'] === '계약' && parseFloat(row['주간 필수 근무시간']) >= 1.6
    );

    const calculateAverage = (employees) => {
        if (!employees.length) return 0;
        const total = employees.reduce((sum, emp) => sum + (parseFloat(emp['근로인정시간']) || 0), 0);
        return total / employees.length;
    };

    const calculateOvertimeAverage = (employees) => {
        if (!employees.length) return 0;
        const total = employees.reduce((sum, emp) => {
            const workTime = parseFloat(emp['근로인정시간']) || 0;
            const requiredTime = parseFloat(emp['주간 필수 근무시간']) || 0;
            return sum + Math.max(0, workTime - requiredTime);
        }, 0);
        return total / employees.length;
    };

    const avgWorkRegular = calculateAverage(regularEmployees);
    const avgWorkContract = calculateAverage(fullTimeContractors);
    const avgOvertimeRegular = calculateOvertimeAverage(regularEmployees);
    const avgOvertimeContract = calculateOvertimeAverage(fullTimeContractors);

    // UI 업데이트
    document.getElementById('totalCount').textContent = `${nonExecutives.length}명`;
    document.getElementById('avgWorkRegular').textContent = `정규: ${formatHours(avgWorkRegular)}`;
    document.getElementById('avgWorkContract').textContent = `계약: ${formatHours(avgWorkContract)}`;
    document.getElementById('avgOvertimeRegular').textContent = `정규: ${formatHours(avgOvertimeRegular)}`;
    document.getElementById('avgOvertimeContract').textContent = `계약: ${formatHours(avgOvertimeContract)}`;
}

// 차트 데이터 생성
function createChartData() {
    const allNonExecutives = currentWeekData.filter(row => row['고용구분'] !== '임원');
    const 사업CIC = allNonExecutives.filter(row =>
        row['소속1'] === '사업' || row['소속1'] === 'CIC'
    );
    const 제품 = allNonExecutives.filter(row => row['소속1'] === '제품');
    const 팔도감 = allNonExecutives.filter(row =>
        row['소속1'] === '팔도감' || row['소속1'] === '팔도감CIC'
    );
    const filteredNonExecutives = getFilteredData().filter(row => row['고용구분'] !== '임원');

    const calculateStatsByType = (employees) => {
        const regular = employees.filter(emp => emp['고용구분'] === '정규');
        const fullTimeContract = employees.filter(emp =>
            emp['고용구분'] === '계약' && parseFloat(emp['주간 필수 근무시간']) >= 1.6
        );

        const calculateAverage = (empList) => {
            if (!empList.length) return { workTime: 0, overtime: 0 };

            const totalWorkTime = empList.reduce((sum, emp) =>
                sum + (parseFloat(emp['근로인정시간']) || 0), 0
            );

            const totalOvertime = empList.reduce((sum, emp) => {
                const workTime = parseFloat(emp['근로인정시간']) || 0;
                const requiredTime = parseFloat(emp['주간 필수 근무시간']) || 0;
                return sum + Math.max(0, workTime - requiredTime);
            }, 0);

            return {
                workTime: convertToHours(totalWorkTime / empList.length),
                overtime: convertToHours(totalOvertime / empList.length)
            };
        };

        return {
            regular: calculateAverage(regular),
            contract: calculateAverage(fullTimeContract)
        };
    };

    const stats = {
        전사: calculateStatsByType(allNonExecutives),
        사업CIC: calculateStatsByType(사업CIC),
        제품: calculateStatsByType(제품),
        팔도감: calculateStatsByType(팔도감),
        선택조직: calculateStatsByType(filteredNonExecutives)
    };

    return [
        { name: '전사', ...stats.전사 },
        { name: '사업+CIC', ...stats.사업CIC },
        { name: '제품', ...stats.제품 },
        { name: '팔도감+팔도감CIC', ...stats.팔도감 },
        { name: '선택조직', ...stats.선택조직 }
    ].map(item => ({
        name: item.name,
        정규직근무: Math.round(item.regular.workTime * 10) / 10,
        정규직초과: Math.round(item.regular.overtime * 10) / 10,
        계약직근무: Math.round(item.contract.workTime * 10) / 10,
        계약직초과: Math.round(item.contract.overtime * 10) / 10
    }));
}

// 차트 업데이트
function updateCharts() {
    const chartData = createChartData();
    const labels = chartData.map(item => item.name);

    // 근무시간 차트
    const workTimeCtx = document.getElementById('workTimeChart').getContext('2d');
    if (workTimeChart) {
        workTimeChart.destroy();
    }
    
    workTimeChart = new Chart(workTimeCtx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: '정규직 근무시간',
                    data: chartData.map(item => item.정규직근무),
                    backgroundColor: '#10B981',
                    borderWidth: 1
                },
                {
                    label: '계약직 근무시간',
                    data: chartData.map(item => item.계약직근무),
                    backgroundColor: '#6B7280',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: false,
                    min: 35,
                    max: 50,
                    title: {
                        display: true,
                        text: '시간'
                    }
                }
            },
            plugins: {
                legend: {
                    position: 'top'
                }
            }
        }
    });

    // 초과근무 차트
    const overtimeCtx = document.getElementById('overtimeChart').getContext('2d');
    if (overtimeChart) {
        overtimeChart.destroy();
    }
    
    overtimeChart = new Chart(overtimeCtx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: '정규직 초과근무',
                    data: chartData.map(item => item.정규직초과),
                    backgroundColor: '#EF4444',
                    borderWidth: 1
                },
                {
                    label: '계약직 초과근무',
                    data: chartData.map(item => item.계약직초과),
                    backgroundColor: '#F59E0B',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    min: 0,
                    max: 8,
                    title: {
                        display: true,
                        text: '시간'
                    }
                }
            },
            plugins: {
                legend: {
                    position: 'top'
                }
            }
        }
    });
}

// 직원 테이블 업데이트
function updateEmployeeTable() {
    const filteredData = getFilteredData();
    let processedData = filteredData
        .filter(emp => emp['고용구분'] !== '임원')
        .map(emp => {
            const workTime = parseFloat(emp['근로인정시간']) || 0;
            const requiredTime = parseFloat(emp['주간 필수 근무시간']) || 0;
            const overtime = Math.max(0, workTime - requiredTime);

            return {
                사번: emp['사번'] || '',
                이름: emp['이름'] || '',
                소속1: emp['소속1'] || '',
                소속2: emp['소속2'] || '',
                소속3: emp['소속3'] || '',
                트라이브: emp['트라이브'] || '',
                소속4: emp['소속4'] || '',
                고용구분: emp['고용구분'] || '',
                일평균근무시간: workTime / 5,
                주간총근무시간: workTime,
                주간초과근무시간: overtime
            };
        });

    // 직원 검색 필터 적용
    if (employeeSearch) {
        processedData = processedData.filter(emp =>
            emp.이름.toLowerCase().includes(employeeSearch.toLowerCase())
        );
    }

    // 직원 필터 적용 (배열 필터 방식)
    Object.entries(employeeFilters).forEach(([key, values]) => {
        if (key.endsWith('_search')) return; // 검색어는 제외
        
        if (values && Array.isArray(values) && values.length > 0) {
            processedData = processedData.filter(emp => {
                const empValue = emp[key] || '';
                return values.some(filterValue => 
                    empValue.toString().toLowerCase().includes(filterValue.toLowerCase())
                );
            });
        }
    });

    // 새로운 정렬 시스템 적용
    if (employeeSortConfig.column) {
        processedData = applyEmployeeSorting(processedData);
    }
    // 기존 정렬 시스템도 유지 (호환성)
    else if (sortConfig.key) {
        processedData.sort((a, b) => {
            let aVal = a[sortConfig.key];
            let bVal = b[sortConfig.key];

            if (['일평균근무시간', '주간총근무시간', '주간초과근무시간'].includes(sortConfig.key)) {
                aVal = typeof aVal === 'number' ? aVal : 0;
                bVal = typeof bVal === 'number' ? bVal : 0;
            } else {
                aVal = (aVal || '').toString().toLowerCase();
                bVal = (bVal || '').toString().toLowerCase();
            }

            if (sortConfig.direction === 'asc') {
                return aVal > bVal ? 1 : -1;
            } else {
                return aVal < bVal ? 1 : -1;
            }
        });
    }

    // 테이블 바디 업데이트
    const tbody = document.getElementById('employeeTableBody');
    tbody.innerHTML = '';

    processedData.forEach((employee, index) => {
        const row = document.createElement('tr');
        row.className = 'table-row hover:bg-gray-50';
        
        const employmentTypeClass = 
            employee.고용구분 === '정규' ? 'bg-green-100 text-green-800' :
            employee.고용구분 === '계약' ? 'bg-blue-100 text-blue-800' :
            'bg-gray-100 text-gray-800';

        row.innerHTML = `
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${employee.사번}">${employee.사번}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900" title="${employee.이름}">${employee.이름}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속1}">${employee.소속1}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속2}">${employee.소속2}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속3}">${employee.소속3}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.트라이브}">${employee.트라이브}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속4}">${employee.소속4}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                <span class="px-2 py-1 rounded-full text-xs ${employmentTypeClass}" title="${employee.고용구분}">
                    ${employee.고용구분}
                </span>
            </td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${formatHours(employee.일평균근무시간)}">${formatHours(employee.일평균근무시간)}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${formatHours(employee.주간총근무시간)}">${formatHours(employee.주간총근무시간)}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${formatHours(employee.주간초과근무시간)}">${formatHours(employee.주간초과근무시간)}</td>
        `;
        tbody.appendChild(row);
    });

    // 카운트 업데이트
    document.getElementById('employeeCountText').textContent = 
        `${selectedWeek ? selectedWeek + ' - ' : ''}총 ${processedData.length}명`;
    document.getElementById('tableFooter').textContent = 
        `${selectedWeek ? selectedWeek + ' - ' : ''}총 ${processedData.length}명의 직원 데이터를 표시하고 있습니다.`;

    // 직원 필터 생성
    updateEmployeeFilters();
    
    // 정렬 아이콘 업데이트
    updateEmployeeSortIcons();
}

// 직원 필터 업데이트 (조직 필터와 동일한 방식)
function updateEmployeeFilters() {
    const container = document.getElementById('employeeFilters');
    const filterKeys = ['소속1', '소속2', '소속3', '트라이브', '소속4', '고용구분'];

    container.innerHTML = '';
    
    filterKeys.forEach(key => {
        // 현재 필터링된 데이터에서 해당 키의 고유 값들 추출
        const uniqueValues = [...new Set(
            getFilteredData()
                .filter(emp => emp['고용구분'] !== '임원')
                .map(emp => emp[key])
                .filter(val => val && val.toString().trim() !== '')
        )].sort();

        const div = document.createElement('div');
        div.innerHTML = `
            <label class="block text-xs font-medium text-gray-600 mb-1">${key}</label>
            <input 
                type="text" 
                placeholder="${key} 검색..." 
                class="employee-filter-search w-full p-1 mb-1 border border-gray-300 rounded text-xs focus:ring-1 focus:ring-blue-500"
                data-key="${key}"
            />
            <div class="employee-filter-options border border-gray-300 rounded p-1 max-h-20 overflow-y-auto bg-white" data-key="${key}">
            </div>
            <div class="employee-filter-count text-xs text-blue-600 mt-1" data-key="${key}" style="display: none;">
                0개 선택됨
            </div>
        `;
        container.appendChild(div);

        updateEmployeeFilterOptions(key, uniqueValues);
    });

    // 검색 이벤트 리스너
    document.querySelectorAll('.employee-filter-search').forEach(input => {
        input.addEventListener('input', (e) => {
            const key = e.target.dataset.key;
            employeeFilters[key + '_search'] = e.target.value;
            
            // 현재 데이터에서 고유 값들 다시 추출
            const uniqueValues = [...new Set(
                getFilteredData()
                    .filter(emp => emp['고용구분'] !== '임원')
                    .map(emp => emp[key])
                    .filter(val => val && val.toString().trim() !== '')
            )].sort();
            
            updateEmployeeFilterOptions(key, uniqueValues);
        });
    });
}

// 직원 필터 옵션 업데이트
function updateEmployeeFilterOptions(key, allOptions) {
    const container = document.querySelector(`.employee-filter-options[data-key="${key}"]`);
    const searchTerm = (employeeFilters[key + '_search'] || '').toLowerCase();
    const filteredOptions = allOptions.filter(option =>
        option.toString().toLowerCase().includes(searchTerm)
    );

    container.innerHTML = '';
    
    if (filteredOptions.length > 0) {
        filteredOptions.forEach(option => {
            const label = document.createElement('label');
            label.className = 'employee-filter-checkbox flex items-center space-x-1 hover:bg-gray-50 p-0.5 rounded cursor-pointer';
            label.innerHTML = `
                <input 
                    type="checkbox" 
                    class="rounded border-gray-300 text-blue-600 w-3 h-3"
                    data-key="${key}"
                    data-value="${option}"
                    ${(employeeFilters[key] || []).includes(option) ? 'checked' : ''}
                />
                <span class="text-xs text-gray-700 truncate">${option}</span>
            `;
            container.appendChild(label);

            // 체크박스 이벤트
            label.querySelector('input').addEventListener('change', (e) => {
                handleEmployeeFilterChange(key, option, e.target.checked);
            });
        });
    } else {
        container.innerHTML = '<div class="text-xs text-gray-500 p-1">검색 결과 없음</div>';
    }

    // 선택된 항목 수 업데이트
    const countElement = document.querySelector(`.employee-filter-count[data-key="${key}"]`);
    const selectedCount = (employeeFilters[key] || []).length;
    if (selectedCount > 0) {
        countElement.textContent = `${selectedCount}개 선택됨`;
        countElement.style.display = 'block';
    } else {
        countElement.style.display = 'none';
    }
}

// 직원 필터 변경 핸들러
function handleEmployeeFilterChange(key, value, isChecked) {
    if (!employeeFilters[key]) {
        employeeFilters[key] = [];
    }

    if (isChecked) {
        if (!employeeFilters[key].includes(value)) {
            employeeFilters[key].push(value);
        }
    } else {
        employeeFilters[key] = employeeFilters[key].filter(v => v !== value);
    }

    updateEmployeeFilterOptions(key, [...new Set(
        getFilteredData()
            .filter(emp => emp['고용구분'] !== '임원')
            .map(emp => emp[key])
            .filter(val => val && val.toString().trim() !== '')
    )].sort());
    
    updateEmployeeTable();
}

// 알럿 시트들 로드
function loadAlertSheets(workbook) {
    alertData = []; // 기존 알럿 데이터 초기화
    
    const alertSheets = [
        { name: '일근무미달', type: 'daily', typeName: '일근무 미달' },
        { name: '주근무미달', type: 'weekly', typeName: '주간근무 미달' },
        { name: '코어타임 미준수', type: 'coretime', typeName: '코어타임 위반' }
    ];

    console.log('알럿 시트 로드 시작...');

    alertSheets.forEach(({ name, type, typeName }) => {
        const sheet = workbook.Sheets[name];
        if (sheet) {
            console.log(`${name} 시트 발견, 데이터 로드 중...`);
            const sheetData = XLSX.utils.sheet_to_json(sheet, { raw: false });
            
            // 알럿 여부가 'O'인 항목들만 필터링
            const alertItems = sheetData.filter(row => {
                const alertFlag = row['알럿 여부'] || row['알럿'] || row['검토대상'] || row['알림'] || '';
                return alertFlag.toString().toLowerCase() === 'o';
            });

            console.log(`${name} 시트에서 ${alertItems.length}건의 알럿 발견`);

            // 알럿 데이터로 변환
            alertItems.forEach(item => {
                alertData.push({
                    type: type,
                    typeName: typeName,
                    employee: item,
                    details: getAlertDetails(item, type),
                    severity: type === 'weekly' ? 'high' : 'medium'
                });
            });
        } else {
            console.log(`${name} 시트를 찾을 수 없습니다`);
        }
    });

            console.log(`총 ${alertData.length}건의 알럿 로드 완료`);
        
        // 근무 미달 데이터도 함께 로드
        loadShortageData(workbook);
    }

// 근무 미달 데이터 로드
function loadShortageData(workbook) {
    // 데이터 초기화
    shortageData = {
        daily: [],
        weekly: [],
        coretime: []
    };

    const shortageSheets = [
        { name: '수정요청리스트(일근무)', key: 'daily' },
        { name: '수정요청리스트(주근무)', key: 'weekly' },
        { name: '수정요청리스트(코어타임)', key: 'coretime' }
    ];

    console.log('근무 미달 데이터 로드 시작...');

    shortageSheets.forEach(({ name, key }) => {
        const sheet = workbook.Sheets[name];
        if (sheet) {
            console.log(`${name} 시트 발견, 데이터 로드 중...`);
            const sheetData = XLSX.utils.sheet_to_json(sheet, { raw: false });
            
            // 알럿 여부가 'O'인 항목들만 필터링
            const shortageItems = sheetData.filter(row => {
                const alertFlag = row['알럿 여부'] || row['알럿'] || row['검토대상'] || row['알림'] || '';
                return alertFlag.toString().toLowerCase() === 'o';
            });

            console.log(`${name} 시트에서 ${shortageItems.length}건의 미달 대상자 발견`);
            console.log(`${name} 시트 전체 데이터 수:`, sheetData.length);
            console.log(`${name} 시트 첫 번째 데이터:`, sheetData[0]);
            console.log(`${name} 시트 필터링된 항목들:`, shortageItems);
            shortageData[key] = shortageItems;
        } else {
            console.log(`${name} 시트를 찾을 수 없습니다`);
        }
    });

    console.log('근무 미달 데이터 로드 완료:', shortageData);
    
    // 첫 번째 탭 활성화
    setTimeout(() => {
        switchShortageTab('daily');
    }, 100);
}

// 근무 미달 섹션 업데이트
function updateShortageSection() {
    const shortageSection = document.getElementById('shortageSection');
    if (!shortageSection) {
        console.log('shortageSection 요소를 찾을 수 없습니다');
        return;
    }

    console.log('근무 미달 섹션 업데이트 시작...');
    console.log('현재 shortageData:', shortageData);

    // 섹션 표시
    shortageSection.style.display = 'block';

    // 각 탭별 카운트 업데이트
    updateShortageCounts();
    
    // 현재 활성 탭 컨텐츠 업데이트
    updateShortageTabContent(currentShortageTab);
    
    // 첫 번째 탭 활성화
    setTimeout(() => {
        switchShortageTab('daily');
    }, 100);
}

// 근무 미달 카운트 업데이트
function updateShortageCounts() {
    const filteredShortageData = getFilteredShortageData();
    
    // 각 탭별 카운트 업데이트
    Object.keys(shortageData).forEach(key => {
        const countElement = document.getElementById(`${key}Count`);
        if (countElement) {
            countElement.textContent = filteredShortageData[key].length;
        }
    });
}

// 조직 필터가 적용된 근무 미달 데이터 반환
function getFilteredShortageData() {
    const filtered = {};
    
    Object.keys(shortageData).forEach(key => {
        filtered[key] = shortageData[key].filter(emp => {
            // 조직 필터 적용
            return Object.entries(filters).every(([filterKey, selectedValues]) => {
                if (!selectedValues || selectedValues.length === 0) return true;
                
                const empValue = emp[filterKey] || '';
                return selectedValues.some(filterValue => 
                    empValue.toString().toLowerCase().includes(filterValue.toLowerCase())
                );
            });
        });
    });
    
    return filtered;
}

// 근무 미달 탭 컨텐츠 업데이트
function updateShortageTabContent(tabKey) {
    const filteredData = getFilteredShortageData()[tabKey] || [];
    const tableHeader = document.getElementById(`${tabKey}TableHeader`);
    const tableBody = document.getElementById(`${tabKey}TableBody`);
    const emptyMessage = document.getElementById('shortageEmptyMessage');

    if (!tableHeader || !tableBody) return;

    if (filteredData.length === 0) {
        tableBody.innerHTML = '';
        if (emptyMessage) {
            emptyMessage.style.display = 'block';
        }
        return;
    }

    if (emptyMessage) {
        emptyMessage.style.display = 'none';
    }

    // 테이블 헤더 생성 (첫 번째 데이터 기준) - 모든 컬럼 표시
    if (filteredData.length > 0) {
        const sampleData = filteredData[0];
        const headers = Object.keys(sampleData);

        tableHeader.innerHTML = headers.map(header => {
            let headerClass = 'px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer hover:bg-gray-100 select-none';
            
            // 코어타임 컬럼은 넓게
            if (header.includes('코어타임') && header.includes('지각기준')) {
                headerClass += ' min-w-40';
            }
            // 소속 관련 컬럼들은 좁게
            else if (header.includes('소속') || header.includes('트라이브')) {
                headerClass += ' w-24 max-w-24';
            }
            
            const sortState = shortageSortConfig[tabKey];
            let sortIcon = '';
            if (sortState.column === header) {
                sortIcon = sortState.direction === 'asc' ? ' ↑' : ' ↓';
            }
            
            return `<th class="${headerClass}" data-column="${header}" onclick="sortShortageTable('${tabKey}', '${header}')">${header}${sortIcon}</th>`;
        }).join('');

        // 데이터 정렬 적용
        const sortedData = applyShortageSorting(filteredData, tabKey);
        
        // 테이블 바디 생성
        tableBody.innerHTML = sortedData.map(emp => {
            const cells = headers.map(header => {
                const value = emp[header] || '';
                const formattedValue = formatCellValue(value, header);
                
                let cellClass = 'px-6 py-4 text-sm text-gray-900';
                
                // 코어타임 컬럼은 넓게
                if (header.includes('코어타임') && header.includes('지각기준')) {
                    cellClass += ' min-w-40';
                }
                // 소속 관련 컬럼들은 좁게, 텍스트 잘림 처리
                else if (header.includes('소속') || header.includes('트라이브')) {
                    cellClass += ' w-24 max-w-24 truncate';
                } else {
                    cellClass += ' whitespace-nowrap';
                }
                
                return `<td class="${cellClass}" title="${formattedValue}">${formattedValue}</td>`;
            }).join('');
            
            return `<tr class="hover:bg-gray-50">${cells}</tr>`;
        }).join('');
    }
}

// 근무 미달 탭 전환
function switchShortageTab(tabKey) {
    currentShortageTab = tabKey;

    // 탭 버튼 스타일 업데이트
    document.querySelectorAll('.shortage-tab').forEach(tab => {
        const isActive = tab.dataset.tab === tabKey;
        if (isActive) {
            tab.className = 'shortage-tab py-2 px-1 border-b-2 border-blue-500 text-blue-600 font-medium text-sm whitespace-nowrap';
        } else {
            tab.className = 'shortage-tab py-2 px-1 border-b-2 border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300 font-medium text-sm whitespace-nowrap';
        }
    });

    // 탭 컨텐츠 표시/숨김
    document.querySelectorAll('.shortage-tab-content').forEach(content => {
        content.style.display = content.dataset.tab === tabKey ? 'block' : 'none';
    });

    // 선택된 탭 컨텐츠 업데이트
    updateShortageTabContent(tabKey);
}

// 셀 값 포맷팅 함수 (이제 엑셀에서 포맷팅된 텍스트를 그대로 사용)
function formatCellValue(value, header) {
    if (value === null || value === undefined) return '';
    
    // 엑셀에서 포맷팅된 텍스트 그대로 반환
    return value.toString();
}

// 근무 미달 테이블 정렬
function sortShortageTable(tabKey, column) {
    const currentSort = shortageSortConfig[tabKey];
    
    // 같은 컬럼을 클릭하면 방향 전환, 다른 컬럼이면 오름차순으로 시작
    if (currentSort.column === column) {
        currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
    } else {
        currentSort.column = column;
        currentSort.direction = 'asc';
    }
    
    // 테이블 업데이트
    updateShortageTabContent(tabKey);
}

// 근무 미달 데이터 정렬 적용
function applyShortageSorting(data, tabKey) {
    const sortConfig = shortageSortConfig[tabKey];
    
    if (!sortConfig.column) {
        return data; // 정렬 설정이 없으면 원본 반환
    }
    
    return [...data].sort((a, b) => {
        let aVal = a[sortConfig.column] || '';
        let bVal = b[sortConfig.column] || '';
        
        // 숫자인지 확인
        const aNum = parseFloat(aVal);
        const bNum = parseFloat(bVal);
        const isNumeric = !isNaN(aNum) && !isNaN(bNum);
        
        let comparison = 0;
        
        if (isNumeric) {
            // 숫자 정렬
            comparison = aNum - bNum;
        } else {
            // 문자열 정렬
            comparison = aVal.toString().toLowerCase().localeCompare(bVal.toString().toLowerCase());
        }
        
        return sortConfig.direction === 'asc' ? comparison : -comparison;
    });
}

// 직원 테이블 정렬
function sortEmployeeTable(column) {
    // 같은 컬럼을 클릭하면 방향 전환, 다른 컬럼이면 오름차순으로 시작
    if (employeeSortConfig.column === column) {
        employeeSortConfig.direction = employeeSortConfig.direction === 'asc' ? 'desc' : 'asc';
    } else {
        employeeSortConfig.column = column;
        employeeSortConfig.direction = 'asc';
    }
    
    // 기존 정렬 설정 초기화 (새로운 정렬 시스템 사용)
    sortConfig.key = null;
    
    // 테이블 업데이트
    updateEmployeeTable();
}

// 직원 데이터 정렬 적용
function applyEmployeeSorting(data) {
    if (!employeeSortConfig.column) {
        return data;
    }
    
    return [...data].sort((a, b) => {
        let aVal = a[employeeSortConfig.column] || '';
        let bVal = b[employeeSortConfig.column] || '';
        
        // 숫자인지 확인
        const aNum = parseFloat(aVal);
        const bNum = parseFloat(bVal);
        const isNumeric = !isNaN(aNum) && !isNaN(bNum);
        
        let comparison = 0;
        
        if (isNumeric) {
            // 숫자 정렬
            comparison = aNum - bNum;
        } else {
            // 문자열 정렬
            comparison = aVal.toString().toLowerCase().localeCompare(bVal.toString().toLowerCase());
        }
        
        return employeeSortConfig.direction === 'asc' ? comparison : -comparison;
    });
}





// 직원 테이블 정렬 아이콘 업데이트
function updateEmployeeSortIcons() {
    // 모든 정렬 아이콘 초기화
    const sortColumns = ['사번', '이름', '소속1', '소속2', '소속3', '트라이브', '소속4', '고용구분', '일평균근무시간', '주간총근무시간', '주간초과근무시간'];
    sortColumns.forEach(col => {
        const element = document.getElementById(`sort-${col}`);
        if (element) element.textContent = '';
    });
    
    // 현재 정렬 아이콘 표시
    if (employeeSortConfig.column) {
        const element = document.getElementById(`sort-${employeeSortConfig.column}`);
        if (element) {
            element.textContent = employeeSortConfig.direction === 'asc' ? ' ↑' : ' ↓';
        }
    }
}

// 초과근무 테이블 업데이트
function updateOvertimeTables() {
    updateRegularOvertimeTable();
    updateContractOvertimeTable();
}

// 정규직 초과근무 테이블 업데이트
function updateRegularOvertimeTable() {
    const filteredData = getFilteredData();
    let processedData = filteredData
        .filter(emp => emp['고용구분'] === '정규') // 정규직만
        .map(emp => {
            const workTime = parseFloat(emp['근로인정시간']) || 0;
            const requiredTime = parseFloat(emp['주간 필수 근무시간']) || 0;
            const overtime = Math.max(0, workTime - requiredTime);

            return {
                사번: emp['사번'] || '',
                이름: emp['이름'] || '',
                소속1: emp['소속1'] || '',
                소속2: emp['소속2'] || '',
                소속3: emp['소속3'] || '',
                트라이브: emp['트라이브'] || '',
                소속4: emp['소속4'] || '',
                고용구분: emp['고용구분'] || '',
                일평균근무시간: workTime / 5,
                주간총근무시간: workTime,
                주간초과근무시간: overtime
            };
        })
        .filter(emp => emp.주간초과근무시간 >= 10/24); // 10시간 이상 (days 단위로 10/24)

    // 정렬 적용
    if (overtimeSortConfig.regular.column) {
        processedData = applyOvertimeSorting(processedData, 'regular');
    }

    // 테이블 바디 업데이트
    const tbody = document.getElementById('regularOvertimeTableBody');
    const section = document.getElementById('regularOvertimeSection');
    
    if (!tbody || !section) return;

    tbody.innerHTML = '';
    section.style.display = 'block';

    if (processedData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="11" class="px-6 py-8 text-center text-gray-500">
                    <div class="flex flex-col items-center">
                        <i data-lucide="users" class="mb-2 text-gray-400" size="32"></i>
                        <p class="text-lg font-medium">주간 초과근무 10시간 이상인 정규직 직원이 없습니다</p>
                        <p class="text-sm text-gray-400 mt-1">모든 정규직 직원이 적정 근무시간을 유지하고 있습니다 👍</p>
                    </div>
                </td>
            </tr>
        `;
        
        // 카운트 업데이트
        document.getElementById('regularOvertimeCountText').textContent = 
            `주간 초과근무 10시간 이상 - 총 0명`;
        document.getElementById('regularOvertimeFooter').textContent = 
            `주간 초과근무 10시간 이상인 정규직 직원이 없습니다.`;
        return;
    }

    processedData.forEach((employee, index) => {
        const row = document.createElement('tr');
        row.className = 'table-row hover:bg-gray-50';
        
        const employmentTypeClass = 'bg-green-100 text-green-800'; // 정규직

        row.innerHTML = `
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${employee.사번}">${employee.사번}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900" title="${employee.이름}">${employee.이름}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속1}">${employee.소속1}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속2}">${employee.소속2}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속3}">${employee.소속3}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.트라이브}">${employee.트라이브}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속4}">${employee.소속4}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                <span class="px-2 py-1 rounded-full text-xs ${employmentTypeClass}" title="${employee.고용구분}">
                    ${employee.고용구분}
                </span>
            </td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${formatHours(employee.일평균근무시간)}">${formatHours(employee.일평균근무시간)}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${formatHours(employee.주간총근무시간)}">${formatHours(employee.주간총근무시간)}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${formatHours(employee.주간초과근무시간)}">${formatHours(employee.주간초과근무시간)}</td>
        `;
        tbody.appendChild(row);
    });

    // 카운트 업데이트
    document.getElementById('regularOvertimeCountText').textContent = 
        `주간 초과근무 10시간 이상 - 총 ${processedData.length}명`;
    document.getElementById('regularOvertimeFooter').textContent = 
        `주간 초과근무 10시간 이상인 정규직 직원 ${processedData.length}명을 표시하고 있습니다.`;
    
    // 정렬 아이콘 업데이트
    updateOvertimeSortIcons('regular');
}

// 계약직 초과근무 테이블 업데이트
function updateContractOvertimeTable() {
    const filteredData = getFilteredData();
    let processedData = filteredData
        .filter(emp => emp['고용구분'] === '계약') // 계약직만
        .map(emp => {
            const workTime = parseFloat(emp['근로인정시간']) || 0;
            const requiredTime = parseFloat(emp['주간 필수 근무시간']) || 0;
            const overtime = Math.max(0, workTime - requiredTime);

            return {
                사번: emp['사번'] || '',
                이름: emp['이름'] || '',
                소속1: emp['소속1'] || '',
                소속2: emp['소속2'] || '',
                소속3: emp['소속3'] || '',
                트라이브: emp['트라이브'] || '',
                소속4: emp['소속4'] || '',
                고용구분: emp['고용구분'] || '',
                일평균근무시간: workTime / 5,
                주간총근무시간: workTime,
                주간초과근무시간: overtime
            };
        })
        .filter(emp => emp.주간초과근무시간 >= 3/24); // 3시간 이상 (days 단위로 3/24)

    // 정렬 적용
    if (overtimeSortConfig.contract.column) {
        processedData = applyOvertimeSorting(processedData, 'contract');
    }

    // 테이블 바디 업데이트
    const tbody = document.getElementById('contractOvertimeTableBody');
    const section = document.getElementById('contractOvertimeSection');
    
    if (!tbody || !section) return;

    tbody.innerHTML = '';
    section.style.display = 'block';

    if (processedData.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="11" class="px-6 py-8 text-center text-gray-500">
                    <div class="flex flex-col items-center">
                        <i data-lucide="clock" class="mb-2 text-gray-400" size="32"></i>
                        <p class="text-lg font-medium">주간 초과근무 3시간 이상인 계약직 직원이 없습니다</p>
                        <p class="text-sm text-gray-400 mt-1">모든 계약직 직원이 적정 근무시간을 유지하고 있습니다 👍</p>
                    </div>
                </td>
            </tr>
        `;
        
        // 카운트 업데이트
        document.getElementById('contractOvertimeCountText').textContent = 
            `주간 초과근무 3시간 이상 - 총 0명`;
        document.getElementById('contractOvertimeFooter').textContent = 
            `주간 초과근무 3시간 이상인 계약직 직원이 없습니다.`;
        return;
    }

    processedData.forEach((employee, index) => {
        const row = document.createElement('tr');
        row.className = 'table-row hover:bg-gray-50';
        
        const employmentTypeClass = 'bg-blue-100 text-blue-800'; // 계약직

        row.innerHTML = `
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${employee.사번}">${employee.사번}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900" title="${employee.이름}">${employee.이름}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속1}">${employee.소속1}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속2}">${employee.소속2}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속3}">${employee.소속3}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.트라이브}">${employee.트라이브}</td>
            <td class="px-6 py-4 text-sm text-gray-500 w-24 max-w-24 truncate" title="${employee.소속4}">${employee.소속4}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                <span class="px-2 py-1 rounded-full text-xs ${employmentTypeClass}" title="${employee.고용구분}">
                    ${employee.고용구분}
                </span>
            </td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${formatHours(employee.일평균근무시간)}">${formatHours(employee.일평균근무시간)}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${formatHours(employee.주간총근무시간)}">${formatHours(employee.주간총근무시간)}</td>
            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900" title="${formatHours(employee.주간초과근무시간)}">${formatHours(employee.주간초과근무시간)}</td>
        `;
        tbody.appendChild(row);
    });

    // 카운트 업데이트
    document.getElementById('contractOvertimeCountText').textContent = 
        `주간 초과근무 3시간 이상 - 총 ${processedData.length}명`;
    document.getElementById('contractOvertimeFooter').textContent = 
        `주간 초과근무 3시간 이상인 계약직 직원 ${processedData.length}명을 표시하고 있습니다.`;
    
    // 정렬 아이콘 업데이트
    updateOvertimeSortIcons('contract');
}

// 초과근무 테이블 정렬
function sortOvertimeTable(type, column) {
    const currentSort = overtimeSortConfig[type];
    
    // 같은 컬럼을 클릭하면 방향 전환, 다른 컬럼이면 오름차순으로 시작
    if (currentSort.column === column) {
        currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
    } else {
        currentSort.column = column;
        currentSort.direction = 'asc';
    }
    
    // 테이블 업데이트
    if (type === 'regular') {
        updateRegularOvertimeTable();
    } else if (type === 'contract') {
        updateContractOvertimeTable();
    }
}

// 초과근무 데이터 정렬 적용
function applyOvertimeSorting(data, type) {
    const sortConfig = overtimeSortConfig[type];
    
    if (!sortConfig.column) {
        return data;
    }
    
    return [...data].sort((a, b) => {
        let aVal = a[sortConfig.column] || '';
        let bVal = b[sortConfig.column] || '';
        
        // 숫자인지 확인
        const aNum = parseFloat(aVal);
        const bNum = parseFloat(bVal);
        const isNumeric = !isNaN(aNum) && !isNaN(bNum);
        
        let comparison = 0;
        
        if (isNumeric) {
            // 숫자 정렬
            comparison = aNum - bNum;
        } else {
            // 문자열 정렬
            comparison = aVal.toString().toLowerCase().localeCompare(bVal.toString().toLowerCase());
        }
        
        return sortConfig.direction === 'asc' ? comparison : -comparison;
    });
}

// 초과근무 테이블 정렬 아이콘 업데이트
function updateOvertimeSortIcons(type) {
    // 모든 정렬 아이콘 초기화
    const sortColumns = ['사번', '이름', '소속1', '소속2', '소속3', '트라이브', '소속4', '고용구분', '일평균근무시간', '주간총근무시간', '주간초과근무시간'];
    sortColumns.forEach(col => {
        const element = document.getElementById(`sort-${type}-${col}`);
        if (element) element.textContent = '';
    });
    
    // 현재 정렬 아이콘 표시
    const sortConfig = overtimeSortConfig[type];
    if (sortConfig.column) {
        const element = document.getElementById(`sort-${type}-${sortConfig.column}`);
        if (element) {
            element.textContent = sortConfig.direction === 'asc' ? ' ↑' : ' ↓';
        }
    }
}

// 알럿 세부사항 생성
function getAlertDetails(emp, type) {
    const workTime = parseFloat(emp['근로인정시간']) || 0;
    const requiredTime = parseFloat(emp['주간 필수 근무시간']) || 0;
    
    switch (type) {
        case 'weekly':
            return `필수: ${formatHours(requiredTime)}, 실제: ${formatHours(workTime)}`;
        case 'daily':
            const dailyAvg = workTime / 5;
            return `일평균: ${formatHours(dailyAvg)}`;
        case 'coretime':
            return emp['위반사유'] || emp['사유'] || '코어타임 미준수';
        default:
            return '세부사항 없음';
    }
}

// 알럿 감지 함수 (기존 로직은 백업용으로 유지)
function detectAlerts() {
    // 이제 엑셀 시트에서 이미 로드된 알럿 데이터를 사용
    return alertData;
}

function updateAlerts() {
    alertData = detectAlerts();
    // 알럿 테이블이 존재할 때만 업데이트
    if (document.getElementById('alertTableBody')) {
        updateAlertTable();
    }
}

function updateAlertTable() {
    // 필터 적용
    filteredAlertData = alertData.filter(alert => {
        // 알럿 타입 필터
        if (!alertFilters[alert.type]) return false;
        
        // 검색 필터
        if (alertSearchTerm) {
            const searchTerm = alertSearchTerm.toLowerCase();
            return (alert.employee.이름 || '').toLowerCase().includes(searchTerm) ||
                   (alert.employee.사번 || '').toLowerCase().includes(searchTerm);
        }
        
        return true;
    });

    // 테이블 바디 업데이트 (안전하게 요소 확인)
    const alertTableBody = document.getElementById('alertTableBody');
    if (!alertTableBody) {
        console.warn('alertTableBody 요소를 찾을 수 없습니다');
        return;
    }
    alertTableBody.innerHTML = '';

    if (filteredAlertData.length === 0) {
        const message = alertData.length === 0 ? 
            '🎉 알럿 대상자가 없습니다!<br><small class="text-gray-400">모든 직원이 정상적으로 근무하고 있습니다.</small>' : 
            '필터된 알럿 대상자가 없습니다.';
        
        alertTableBody.innerHTML = `
            <tr>
                <td colspan="8" class="px-6 py-8 text-center text-gray-500">
                    ${message}
                </td>
            </tr>
        `;
    } else {
        filteredAlertData.forEach((alert, index) => {
            const row = document.createElement('tr');
            row.className = 'table-row hover:bg-gray-50';
            
            const alertTypeClass = 
                alert.type === 'weekly' ? 'bg-red-100 text-red-800' :
                alert.type === 'coretime' ? 'bg-yellow-100 text-yellow-800' :
                'bg-orange-100 text-orange-800';

            const emp = alert.employee;
            const workTime = parseFloat(emp['근로인정시간']) || 0;

            row.innerHTML = `
                <td class="px-6 py-4 whitespace-nowrap text-sm">
                    <span class="px-2 py-1 rounded-full text-xs ${alertTypeClass}">
                        ${alert.typeName}
                    </span>
                </td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${emp.사번 || '-'}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${emp.이름 || '-'}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${emp.소속1 || '-'}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${emp.소속2 || '-'}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${emp.트라이브 || '-'}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${alert.details}</td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${formatHours(workTime)}</td>
            `;
            alertTableBody.appendChild(row);
        });
    }

    // 카운트 업데이트 (안전하게 요소 확인)
    const weeklyCnt = alertData.filter(a => a.type === 'weekly').length;
    const coretimeCnt = alertData.filter(a => a.type === 'coretime').length;
    const dailyCnt = alertData.filter(a => a.type === 'daily').length;

    const weeklyCountEl = document.getElementById('weeklyCount');
    const overtimeCountEl = document.getElementById('overtimeCount');
    const dailyCountEl = document.getElementById('dailyCount');
    const alertCountTextEl = document.getElementById('alertCountText');
    const alertFooterEl = document.getElementById('alertFooter');

    if (weeklyCountEl) weeklyCountEl.textContent = weeklyCnt;
    if (overtimeCountEl) overtimeCountEl.textContent = coretimeCnt;
    if (dailyCountEl) dailyCountEl.textContent = dailyCnt;
    if (alertCountTextEl) alertCountTextEl.textContent = `총 ${filteredAlertData.length}명의 알럿 대상자`;
    if (alertFooterEl) alertFooterEl.textContent = 
        `총 ${filteredAlertData.length}명의 알럿 대상자가 표시되고 있습니다. (전체: ${alertData.length}명)`;
}

function exportAlerts() {
    if (filteredAlertData.length === 0) {
        showError('내보낼 알럿 데이터가 없습니다.');
        return;
    }

    // CSV 형태로 데이터 변환
    const headers = ['알럿타입', '사번', '이름', '소속1', '소속2', '트라이브', '위반세부사항', '현재근무시간'];
    const csvContent = [
        headers.join(','),
        ...filteredAlertData.map(alert => {
            const emp = alert.employee;
            const workTime = parseFloat(emp['근로인정시간']) || 0;
            return [
                alert.typeName,
                emp.사번 || '',
                emp.이름 || '',
                emp.소속1 || '',
                emp.소속2 || '',
                emp.트라이브 || '',
                alert.details,
                formatHours(workTime)
            ].join(',');
        })
    ].join('\n');

    // 파일 다운로드
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', `근무수정알럿_${selectedWeek || new Date().toISOString().slice(0, 10)}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showSuccessMessage('알럿 데이터를 CSV 파일로 내보냈습니다.');
}

// 드래그 앤 드롭 기능
function setupDragAndDrop() {
    const fileUploadArea = document.querySelector('.file-upload-area');

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        fileUploadArea.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        fileUploadArea.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        fileUploadArea.addEventListener(eventName, unhighlight, false);
    });

    function highlight(e) {
        fileUploadArea.style.borderColor = '#3b82f6';
        fileUploadArea.style.backgroundColor = '#f8fafc';
    }

    function unhighlight(e) {
        fileUploadArea.style.borderColor = '#d1d5db';
        fileUploadArea.style.backgroundColor = '';
    }

    fileUploadArea.addEventListener('drop', handleDrop, false);

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;

        if (files.length > 0) {
            fileInput.files = files;
            handleFileUpload({ target: { files: files } });
        }
    }
}

// 이벤트 리스너들
document.addEventListener('DOMContentLoaded', () => {
    // DOM 요소 확인 및 파일 업로드 이벤트
    const fileInputElement = document.getElementById('fileInput');
    if (fileInputElement) {
        fileInputElement.addEventListener('change', handleFileUpload);
        console.log('파일 업로드 이벤트 리스너 연결됨');
    } else {
        console.error('fileInput 요소를 찾을 수 없습니다');
    }

    // 주차 검색 이벤트
    const weekSearchElement = document.getElementById('weekSearch');
    if (weekSearchElement) {
        weekSearchElement.addEventListener('input', (e) => {
            weekSearchTerm = e.target.value;
            updateWeekList();
        });
    }

    // 주차 삭제 버튼 이벤트
    const deleteSelectedWeekBtn = document.getElementById('deleteSelectedWeek');
    if (deleteSelectedWeekBtn) {
        deleteSelectedWeekBtn.addEventListener('click', deleteSelectedWeek);
    }

    // 관리자 설정 모달 이벤트
    const adminToggle = document.getElementById('adminToggle');
    const adminModal = document.getElementById('adminModal');
    const cancelAdmin = document.getElementById('cancelAdmin');
    const saveAdmin = document.getElementById('saveAdmin');

    if (adminToggle) {
        adminToggle.addEventListener('click', () => {
            // 기존 설정 표시
            document.getElementById('githubUsername').value = githubConfig.username || '';
            document.getElementById('githubRepo').value = githubConfig.repo || '';
            document.getElementById('githubToken').value = githubConfig.token || '';
            
            adminModal.style.display = 'flex';
        });
    }

    if (cancelAdmin) {
        cancelAdmin.addEventListener('click', () => {
            adminModal.style.display = 'none';
        });
    }

    if (saveAdmin) {
        saveAdmin.addEventListener('click', async () => {
            const username = document.getElementById('githubUsername').value.trim();
            const repo = document.getElementById('githubRepo').value.trim();
            const token = document.getElementById('githubToken').value.trim();

            if (!username || !repo || !token) {
                showError('모든 필드를 입력해주세요.');
                return;
            }

            try {
                // GitHub 설정 저장
                saveGithubConfig(username, repo, token);
                toggleAdminMode(true);
                
                adminModal.style.display = 'none';
                showSuccessMessage('GitHub 설정이 저장되었습니다. 관리자 모드가 활성화되었습니다.');
                
                // GitHub에서 데이터 다시 로드
                setLoading(true);
                const githubData = await loadFromGithub();
                
                // 기존 데이터와 병합
                Object.keys(githubData).forEach(weekName => {
                    weeklyDataSets[weekName] = githubData[weekName].data;
                });
                
                availableWeeks = Object.keys(weeklyDataSets).sort().reverse();
                updateWeekOptions();
                
                if (availableWeeks.length > 0) {
                    handleWeekSelection(availableWeeks[0]);
                }
                
            } catch (error) {
                showError(`GitHub 설정 오류: ${error.message}`);
            } finally {
                setLoading(false);
            }
        });
    }

    // 모달 외부 클릭 시 닫기
    if (adminModal) {
        adminModal.addEventListener('click', (e) => {
            if (e.target === adminModal) {
                adminModal.style.display = 'none';
            }
        });
    }

    // 직원 검색 이벤트
    document.getElementById('employeeSearch').addEventListener('input', (e) => {
        employeeSearch = e.target.value;
        updateEmployeeTable();
    });

    // 직원 필터 초기화 버튼
    const clearEmployeeFiltersBtn = document.getElementById('clearEmployeeFilters');
    if (clearEmployeeFiltersBtn) {
        clearEmployeeFiltersBtn.addEventListener('click', () => {
            // 필터 초기화
            employeeFilters = {
                소속1: [], 소속2: [], 소속3: [], 트라이브: [], 소속4: [], 고용구분: []
            };
            
            // 검색어 초기화
            document.querySelectorAll('.employee-filter-search').forEach(input => {
                input.value = '';
            });
            
            // 필터 UI 다시 생성
            updateEmployeeFilters();
            updateEmployeeTable();
        });
    }

    // 알럿 관련 이벤트 리스너들 (안전하게 요소 확인 후 연결)
    const refreshAlertsBtn = document.getElementById('refreshAlerts');
    const exportAlertsBtn = document.getElementById('exportAlerts');
    const alertSearchInput = document.getElementById('alertSearch');
    const alertWeeklyCheckbox = document.getElementById('alertWeekly');
    const alertOvertimeCheckbox = document.getElementById('alertOvertime');
    const alertDailyCheckbox = document.getElementById('alertDaily');

    if (refreshAlertsBtn) {
        refreshAlertsBtn.addEventListener('click', () => {
            updateAlerts();
            showSuccessMessage('알럿 대상자를 새로고침했습니다.');
        });
    }

    if (exportAlertsBtn) {
        exportAlertsBtn.addEventListener('click', exportAlerts);
    }

    if (alertSearchInput) {
        alertSearchInput.addEventListener('input', (e) => {
            alertSearchTerm = e.target.value;
            updateAlertTable();
        });
    }

    if (alertWeeklyCheckbox) {
        alertWeeklyCheckbox.addEventListener('change', (e) => {
            alertFilters.weekly = e.target.checked;
            updateAlertTable();
        });
    }

    if (alertOvertimeCheckbox) {
        alertOvertimeCheckbox.addEventListener('change', (e) => {
            alertFilters.coretime = e.target.checked;
            updateAlertTable();
        });
    }

    if (alertDailyCheckbox) {
        alertDailyCheckbox.addEventListener('change', (e) => {
            alertFilters.daily = e.target.checked;
            updateAlertTable();
        });
    }

    // 테이블 정렬 이벤트
    document.addEventListener('click', (e) => {
        if (e.target.dataset.sort) {
            const key = e.target.dataset.sort;
            
            if (sortConfig.key === key) {
                sortConfig.direction = sortConfig.direction === 'asc' ? 'desc' : 'asc';
            } else {
                sortConfig.key = key;
                sortConfig.direction = 'asc';
            }

            // 정렬 표시 업데이트
            document.querySelectorAll('.sort-indicator').forEach(el => el.textContent = '');
            e.target.querySelector('.sort-indicator').textContent = 
                sortConfig.direction === 'asc' ? '↑' : '↓';

            updateEmployeeTable();
        }
    });

    // 드래그 앤 드롭 설정
    setupDragAndDrop();
    
    // 앱 초기화
    initializeApp();
    
    // Lucide 아이콘 초기화
    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }
});
