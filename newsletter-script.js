/**
 * 네이버 API와 NewsAPI를 활용한 통합 뉴스레터 시스템
 * 국내 뉴스는 네이버 API, 해외 뉴스는 NewsAPI로 검색
 * Gemini AI로 관련성을 평가하여 맞춤형 뉴스레터 발송
 * 관련성 7.0 이상 뉴스만 포함, 상세 로깅 추가
 */

// ---------------------- 상수 및 설정 ----------------------

// API 및 데이터 관련 상수
const CONFIG = {
  DEFAULT_EMAIL: "woomir@gmail.com",
  DEFAULT_GEMINI_MODEL: "gemini-2.0-flash",
  GEMINI_MODEL_PROPERTY: "GEMINI_MODEL",
  BATCH_SIZE: 5,
  MAX_BATCHES: 5,
  MAX_NEWS_RESULTS: 10,
  MAX_RETRIES: 3,
  DEFAULT_LOCALE: {
    KOREAN: { language: "ko", country: "KR" },
    ENGLISH: { language: "en", country: "US" }
  },
  RELEVANCE_THRESHOLD: {
    DOMESTIC: 7.0,  // 국내 뉴스 필터링 임계값 (7.0 이상만 포함)
    GLOBAL: 7.0,    // 해외 뉴스 필터링 임계값 (7.0 이상만 포함)
    FALLBACK_DOMESTIC: 7.0,
    FALLBACK_GLOBAL: 7.0
  },
  NEWSAPI: {
    ENDPOINT: "https://newsapi.org/v2/",
    ARTICLES_PER_REQUEST: 30,     // 한 번에 가져올 기사 수
    SORT_BY: "publishedAt"
  },
  NAVER_API: {
    ENDPOINT: "https://openapi.naver.com/v1/search/news.json",
    DISPLAY: 30,    // 한 번에 검색할 뉴스 수
    SORT: "date"    // 최신순 정렬
  }
};

// 한글 요일 변환 배열
const KOREAN_DAYS = ["일", "월", "화", "수", "목", "금", "토"];

// ---------------------- 메인 함수 ----------------------

/**
 * 스프레드시트에서 정보를 가져와 뉴스레터 전송을 시작하는 메인 함수
 */
function sendNewsletterEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = Math.max(sheet.getLastRow(), 10);
  
  // 네이버 API 인증 정보 가져오기
  const naverClientId = sheet.getRange("C2").getValue();
  const naverClientSecret = sheet.getRange("D2").getValue();
  
  // 네이버 API 인증 정보 확인
  if (!naverClientId || !naverClientSecret) {
    Logger.log("네이버 API 인증 정보가 없습니다. C2와 D2 셀을 확인해주세요.");
    return;
  }
  
  // 주제 및 옵션 수집
  const topics = collectTopics(sheet, lastRow);
  
  // 이메일 주소 수집
  const emails = collectEmails(sheet, lastRow);
  
  // Google AI API 키 가져오기
  const googleApiKey = getGoogleApiKey(sheet, lastRow);
  
  // 실행 검증
  if (topics.length === 0) {
    Logger.log("검색할 주제가 없습니다.");
    return;
  }
  
  if (!googleApiKey) {
    Logger.log("Google AI API 키가 없습니다.");
    return;
  }
  
  // 뉴스레터 생성 및 발송 (Gemini 중복 검사 적용)
  createAndSendNewsletterWithGeminiCheck(topics, emails, googleApiKey, naverClientId, naverClientSecret);
}

/**
 * 스프레드시트에서 주제 및 관련 설정 수집
 * @param {SpreadsheetApp.Sheet} sheet - 활성 시트
 * @param {number} lastRow - 마지막 행 번호
 * @return {Array} 주제 객체 배열
 */
function collectTopics(sheet, lastRow) {
  const topics = [];
  
  // 먼저 공통 NewsAPI 키 찾기
  let commonNewsApiKey = "";
  for (let i = 2; i <= lastRow; i++) {
    const apiKey = sheet.getRange(`J${i}`).getValue();
    if (apiKey && apiKey.trim() !== "") {
      commonNewsApiKey = apiKey;
      break;
    }
  }
  
  for (let i = 2; i <= lastRow; i++) {
    const topic = sheet.getRange(`A${i}`).getValue();
    if (topic && topic.trim() !== "") {
      const relatedConcepts = sheet.getRange(`F${i}`).getValue();
      const includeGlobalNews = (sheet.getRange(`G${i}`).getValue() || "N").toUpperCase() === "Y";
      const englishKeyword = sheet.getRange(`I${i}`).getValue();
      // 개별 행의 API 키를 확인하고, 없으면 공통 키 사용
      const rowNewsApiKey = sheet.getRange(`J${i}`).getValue();
      const newsApiKey = (rowNewsApiKey && rowNewsApiKey.trim() !== "") ? rowNewsApiKey : commonNewsApiKey;
      
      topics.push({
        topic: topic,
        relatedConcepts: relatedConcepts || "",
        includeGlobalNews: includeGlobalNews,
        englishKeyword: englishKeyword || "",
        newsApiKey: newsApiKey
      });
    }
  }
  
  return topics;
}

/**
 * 스프레드시트에서 이메일 주소 수집
 * @param {SpreadsheetApp.Sheet} sheet - 활성 시트
 * @param {number} lastRow - 마지막 행 번호
 * @return {Array} 이메일 주소 배열
 */
function collectEmails(sheet, lastRow) {
  const emails = [];
  
  for (let i = 2; i <= lastRow; i++) {
    const email = sheet.getRange(`B${i}`).getValue();
    if (email && email.trim() !== "" && email.includes("@") && !emails.includes(email)) {
      emails.push(email);
    }
  }
  
  // 기본 이메일이 없으면 설정
  if (emails.length === 0) {
    emails.push(CONFIG.DEFAULT_EMAIL);
  }
  
  return emails;
}

/**
 * Google AI API 키 가져오기
 * @param {SpreadsheetApp.Sheet} sheet - 활성 시트
 * @param {number} lastRow - 마지막 행 번호
 * @return {string} API 키
 */
function getGoogleApiKey(sheet, lastRow) {
  for (let i = 2; i <= lastRow; i++) {
    const apiKey = sheet.getRange(`E${i}`).getValue();
    if (apiKey && apiKey.trim() !== "") {
      return apiKey;
    }
  }
  return "";
}

// ---------------------- 뉴스레터 생성 및 발송 ----------------------

/**
 * NewsAPI 및 네이버 API 검색 및 이메일 생성/발송 함수
 * @param {Array} topics - 주제 객체 배열
 * @param {Array} emails - 이메일 주소 배열
 * @param {string} googleApiKey - Google AI API 키
 * @param {string} naverClientId - 네이버 API Client ID
 * @param {string} naverClientSecret - 네이버 API Client Secret
 */
function createAndSendNewsletter(topics, emails, googleApiKey, naverClientId, naverClientSecret) {
  // 날짜 정보 설정
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  
  const dateInfo = formatDateForNewsletter(today);
  
  // 모든 주제를 하나의 문자열로 결합 (이메일 제목용)
  const topicNames = topics.map(t => t.topic);
  const mainTopic = topicNames.length > 0 ? `${topicNames[0]}/${topicNames.slice(1).join('/')}` : "맞춤 주제";
  
  // 이메일 본문 초기화
  let emailBody = createNewsletterHeader(mainTopic, dateInfo);
  
  // 각 주제별 뉴스 검색 및 추가
  let categoryIndex = 0;
  for (const topicObj of topics) {
    // 뉴스 카테고리 헤더 생성
    if (categoryIndex > 0) {
      emailBody += `<hr style="border: 0; height: 1px; background-color: #ddd; margin: 25px 0;">`;
    }
    
    emailBody += `<h3 style="color: #1a73e8; margin-top: 20px; margin-bottom: 15px;">■ ${topicObj.topic}</h3>`;
    categoryIndex++;
    
    // 뉴스 수집 및 처리
    const newsContent = processTopicNewsWithSeparation(topicObj, yesterday, googleApiKey, naverClientId, naverClientSecret);
    emailBody += newsContent || `<p style="color: #666;">관련 뉴스를 찾을 수 없습니다.</p>`;
  }
  
  // 이메일 닫기 태그
  emailBody += `</div>`;
  
  // 각 이메일 주소로 뉴스레터 발송
  sendEmailToRecipients(emails, mainTopic, dateInfo, emailBody);
}

/**
 * 뉴스레터 날짜 형식화
 * @param {Date} date - 날짜 객체
 * @return {string} 형식화된 날짜 문자열
 */
function formatDateForNewsletter(date) {
  const monthDay = Utilities.formatDate(date, "Asia/Seoul", "M월 d일");
  const koreanDayOfWeek = KOREAN_DAYS[date.getDay()];
  return `${monthDay}, ${koreanDayOfWeek}요일`;
}

/**
 * 뉴스레터 헤더 HTML 생성
 * @param {string} mainTopic - 주 주제
 * @param {string} dateString - 형식화된 날짜 문자열
 * @return {string} 헤더 HTML
 */
function createNewsletterHeader(mainTopic, dateString) {
  let header = `<div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; color: #333;">`;
  header += `<h2 style="color: #1a73e8; margin-bottom: 5px;">${mainTopic} 뉴스 업데이트</h2>`;
  header += `<p style="color: #666; font-size: 14px; margin-top: 0;">(${dateString})</p>`;
  return header;
}

/**
 * 주제별 뉴스 처리 및 HTML 생성 (국내/해외 뉴스 구분)
 * @param {Object} topicObj - 주제 객체
 * @param {Date} yesterday - 어제 날짜
 * @param {string} googleApiKey - Google AI API 키
 * @param {string} naverClientId - 네이버 API Client ID
 * @param {string} naverClientSecret - 네이버 API Client Secret
 * @return {string} 뉴스 HTML 콘텐츠
 */
function processTopicNewsWithSeparation(topicObj, yesterday, googleApiKey, naverClientId, naverClientSecret) {
  // 주제 정보 추출
  const { topic, relatedConcepts, includeGlobalNews, englishKeyword, newsApiKey } = topicObj;
  
  // 국내 뉴스와 해외 뉴스 배열 초기화
  let domesticNews = [];
  let globalNews = [];
  
  // 국내 뉴스 검색 및 처리 (네이버 API 사용)
  const koreanNewsItems = searchNaverNews(topic, naverClientId, naverClientSecret);
  if (koreanNewsItems.length > 0) {
    Logger.log(`'${topic}' 주제로 네이버에서 검색된 국내 뉴스: ${koreanNewsItems.length}개`);
    
    // AI 분석 및 필터링
    const filteredLocalNews = filterNewsByGeminiBatched(koreanNewsItems, topic, relatedConcepts, googleApiKey, "국내 뉴스");
    Logger.log(`'${topic}' 주제에 대한 Gemini 분석 후 관련성 높은 국내 뉴스: ${filteredLocalNews.length}개`);
    
    // 국내 뉴스 추가
    if (filteredLocalNews.length > 0) {
      domesticNews = prepareNewsItems(filteredLocalNews, "국내");
    }
  }
  
  // 해외 뉴스 검색 및 처리 (NewsAPI 사용)
  if (includeGlobalNews && newsApiKey) {
    const searchKeyword = englishKeyword || topic;
    const englishNewsItems = searchNewsAPI(searchKeyword, "en", newsApiKey);
    
    if (englishNewsItems.length > 0) {
      Logger.log(`'${topic}' 주제로 NewsAPI에서 해외 뉴스 ${englishNewsItems.length}개 검색됨 (검색 키워드: ${searchKeyword})`);
      
      // AI 분석 및 필터링
      const filteredGlobalNews = filterNewsByGeminiBatched(englishNewsItems, topic, relatedConcepts, googleApiKey, "해외 뉴스");
      Logger.log(`'${topic}' 주제에 대한 Gemini 분석 후 관련성 높은 해외 뉴스: ${filteredGlobalNews.length}개`);
      
      // 해외 뉴스 번역 및 추가
      if (filteredGlobalNews.length > 0) {
        const translatedNews = translateNewsWithGemini(filteredGlobalNews, googleApiKey);
        globalNews = prepareNewsItems(translatedNews, "해외");
      }
    }
  } else if (includeGlobalNews && !newsApiKey) {
    Logger.log(`'${topic}' 주제의 해외 뉴스 검색을 위한 NewsAPI 키가 없습니다.`);
  }
  
  // 각 카테고리별로 관련성 점수로 정렬
  domesticNews.sort((a, b) => b.aiRelevanceScore - a.aiRelevanceScore);
  globalNews.sort((a, b) => b.aiRelevanceScore - a.aiRelevanceScore);
  
  Logger.log(`'${topic}' 주제에 대한 총 관련성 높은 국내 뉴스: ${domesticNews.length}개`);
  Logger.log(`'${topic}' 주제에 대한 총 관련성 높은 해외 뉴스: ${globalNews.length}개`);
  
  // 뉴스 HTML 생성 (카테고리별로 구분)
  let newsHtml = '';
  
  // 국내 뉴스 섹션
  if (domesticNews.length > 0) {
    newsHtml += `<h4 style="color: #555; margin-top: 15px; margin-bottom: 10px; border-bottom: 1px solid #eee; padding-bottom: 5px;">🇰🇷 국내 뉴스 (${domesticNews.length}개)</h4>`;
    newsHtml += createNewsItemsHtml(domesticNews, topic, yesterday, googleApiKey);
  }
  
  // 해외 뉴스 섹션
  if (globalNews.length > 0) {
    newsHtml += `<h4 style="color: #555; margin-top: 20px; margin-bottom: 10px; border-bottom: 1px solid #eee; padding-bottom: 5px;">🌏 해외 뉴스 (${globalNews.length}개)</h4>`;
    newsHtml += createNewsItemsHtml(globalNews, topic, yesterday, googleApiKey);
  }
  
  // 뉴스가 없는 경우
  if (domesticNews.length === 0 && globalNews.length === 0) {
    newsHtml += `<p style="color: #666;">관련 뉴스를 찾을 수 없습니다.</p>`;
  }
  
  return newsHtml;
}

/**
 * 뉴스 아이템 준비 (소스 추가)
 * @param {Array} newsItems - 뉴스 아이템 배열
 * @param {string} type - 뉴스 유형 (국내/해외)
 * @return {Array} 소스 정보가 추가된 뉴스 아이템 배열
 */
function prepareNewsItems(newsItems, type) {
  return newsItems.map(item => ({
    ...item,
    type: type
  }));
}

/**
 * 뉴스 HTML 콘텐츠 생성 - 국내 뉴스는 source를 표시하지 않도록 수정
 * @param {Array} newsItems - 뉴스 아이템 배열
 * @param {string} topic - 주제
 * @param {Date} yesterday - 어제 날짜
 * @param {string} googleApiKey - Google API 키
 * @return {string} 뉴스 HTML 콘텐츠
 */
function createNewsItemsHtml(newsItems, topic, yesterday, googleApiKey) {
  let newsHtml = '';
  
  if (newsItems.length > 0) {
    for (const news of newsItems) {
      // AI 기반 뉴스 타입 라벨 결정
      const newsTypeLabel = news.newsType || " ";
      
      // 날짜 형식화
      const pubDateFormatted = news.pubDate ? 
        Utilities.formatDate(news.pubDate, "Asia/Seoul", "M월 d일") : 
        Utilities.formatDate(yesterday, "Asia/Seoul", "M월 d일");
      
      // 요약 생성
      const summary = news.aiSummary || news.description || `${topic}과 관련된 ${news.title} 관련 소식입니다.`;
      
      // 관련성 점수 (10점 만점에 소수점 한 자리까지 표시)
      const relevanceScore = news.aiRelevanceScore ? news.aiRelevanceScore.toFixed(1) : "?";
      
      // 뉴스 종류에 따른 아이콘 (국내/해외)
      const newsTypeIcon = news.type === "국내" ? "🇰🇷" : "🌏";
      
      // 뉴스 아이템 포맷 - 개선된 스타일
      newsHtml += `<div style="margin-bottom: 22px; border-left: 3px solid #1a73e8; padding-left: 12px;">`;
      newsHtml += `<p style="font-weight: bold; margin-bottom: 5px; font-size: 16px;"><strong>(${newsTypeLabel})</strong> ${news.title}</p>`;
      newsHtml += `<p style="margin-top: 0; margin-bottom: 5px; color: #444; font-size: 14px;">${summary}</p>`;
      newsHtml += `<p style="margin-top: 0; font-size: 12px; color: #666;">`;
      
      // 국내 뉴스는 언론사(source) 정보를 표시하지 않음
      if (news.type === "국내") {
        newsHtml += `[${pubDateFormatted}] [관련성:${relevanceScore}/10]<br>`;
      } else {
        // 해외 뉴스는 기존대로 언론사 정보 표시
        newsHtml += `[${pubDateFormatted}/${news.source}] [관련성:${relevanceScore}/10]<br>`;
      }
      
      newsHtml += `<a href="${news.link}" style="color: #1a73e8; text-decoration: none; font-weight: bold;">바로가기</a>`;
      newsHtml += `</p>`;
      newsHtml += `</div>`;
    }
  }
  
  return newsHtml;
}

/**
 * 수신자에게 이메일 발송
 * @param {Array} emails - 이메일 주소 배열
 * @param {string} mainTopic - 주 주제
 * @param {string} dateString - 날짜 문자열
 * @param {string} emailBody - 이메일 HTML 본문
 */
function sendEmailToRecipients(emails, mainTopic, dateString, emailBody) {
  for (const email of emails) {
    // 이메일 발송
    MailApp.sendEmail({
      to: email,
      subject: `[${mainTopic} 뉴스레터] ${dateString} 업데이트`,
      htmlBody: emailBody
    });
    
    Logger.log(`뉴스레터가 ${email}로 발송되었습니다.`);
  }
}

// ---------------------- 네이버 뉴스 검색 API ----------------------

/**
 * 네이버 API를 사용하여 뉴스 검색
 * @param {string} keyword - 검색 키워드
 * @param {string} clientId - 네이버 API Client ID
 * @param {string} clientSecret - 네이버 API Client Secret
 * @return {Array} 뉴스 아이템 배열
 */
function searchNaverNews(keyword, clientId, clientSecret) {
  try {
    // 어제 날짜 계산
    const today = new Date();
    const yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    
    // 키워드 인코딩
    const encodedKeyword = encodeURIComponent(keyword);
    
    // API URL 생성
    const apiUrl = `${CONFIG.NAVER_API.ENDPOINT}?query=${encodedKeyword}&display=${CONFIG.NAVER_API.DISPLAY}&sort=${CONFIG.NAVER_API.SORT}`;
    
    Logger.log(`네이버 뉴스 API 요청 URL: ${apiUrl}`);
    
    // API 요청 헤더 설정
    const headers = {
      "X-Naver-Client-Id": clientId,
      "X-Naver-Client-Secret": clientSecret
    };
    
    // API 요청
    const response = UrlFetchApp.fetch(apiUrl, { 
      headers: headers,
      muteHttpExceptions: true 
    });
    
    // 응답 상태 코드 확인
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      Logger.log(`네이버 API 요청 실패 - 상태 코드: ${responseCode}, 응답: ${response.getContentText().substring(0, 200)}...`);
      return [];
    }
    
    // JSON 응답 파싱
    const jsonResponse = JSON.parse(response.getContentText());
    
    // 뉴스 아이템이 없는 경우
    if (!jsonResponse.items || jsonResponse.items.length === 0) {
      Logger.log(`네이버 API 검색 결과 없음: ${keyword}`);
      return [];
    }
    
    Logger.log(`네이버 API 검색 결과: ${jsonResponse.items.length}개 기사 찾음`);
    
    // 어제 날짜 이후의 뉴스만 필터링
    const yesterdayStart = new Date(yesterday);
    yesterdayStart.setHours(0, 0, 0, 0);
    
    // 응답 변환 및 필터링
    const newsItems = parseNaverNewsResults(jsonResponse.items);
    
    // 날짜로 필터링
    const filteredItems = newsItems.filter(item => 
      item.pubDate >= yesterdayStart && item.pubDate <= today
    );
    
    Logger.log(`네이버 API 검색 결과 중 어제 이후 뉴스: ${filteredItems.length}개`);
    
    return filteredItems;
    
  } catch (error) {
    Logger.log(`네이버 뉴스 API 검색 중 오류 발생: ${error.message}`);
    Logger.log(`스택: ${error.stack}`);
    return [];
  }
}

/**
 * 네이버 API 결과를 내부 형식으로 변환
 * @param {Array} items - 네이버 API 응답 아이템 배열
 * @return {Array} 변환된 뉴스 아이템 배열
 */
function parseNaverNewsResults(items) {
  const newsItems = [];
  
  for (const item of items) {
    try {
      // 필수 필드 확인
      if (!item.title || !item.link) {
        continue;
      }
      
      // HTML 태그 제거
      const title = removeHtmlTags(item.title);
      const description = removeHtmlTags(item.description);
      
      // 발행일 파싱 (네이버는 pubDate 포맷: EEE, dd MMM yyyy HH:mm:ss Z)
      let pubDate = new Date();
      if (item.pubDate) {
        pubDate = new Date(item.pubDate);
      }
      
      // 뉴스 소스 추출 (네이버는 출처를 따로 제공하지 않으므로 링크에서 추출)
      let source = "네이버 뉴스";
      try {
        // 링크에서 도메인 추출 시도
        const urlObj = new URL(item.originallink || item.link);
        source = urlObj.hostname.replace(/^www\./, '');
      } catch (e) {
        // URL 파싱 오류 - 기본값 사용
      }
      
      // 뉴스 아이템 추가
      newsItems.push({
        title: title,
        link: item.link,
        description: description,
        pubDate: pubDate,
        pubDateStr: Utilities.formatDate(pubDate, "Asia/Seoul", "yyyyMMdd"),
        source: source,
        sourceUrl: item.originallink || "",
        isGlobal: false,
        imageUrl: null,
        author: ""
      });
    } catch (itemError) {
      Logger.log(`네이버 뉴스 아이템 파싱 오류: ${itemError.message}`);
    }
  }
  
  return newsItems;
}

/**
 * HTML 태그 제거 헬퍼 함수
 * @param {string} text - 텍스트
 * @return {string} HTML 태그가 제거된 텍스트
 */
function removeHtmlTags(text) {
  if (!text) return "";
  return text.replace(/<[^>]*>/g, "");
}

// ---------------------- NewsAPI 검색 및 필터링 ----------------------

/**
 * NewsAPI를 사용하여 뉴스 검색 (어제 날짜 뉴스만)
 * @param {string} keyword - 검색 키워드
 * @param {string} language - 언어 코드 ('en', 'ko' 등)
 * @param {string} apiKey - NewsAPI 키
 * @return {Array} 뉴스 아이템 배열
 */
function searchNewsAPI(keyword, language = "ko", apiKey) {
  try {
    // 어제 날짜 계산
    const today = new Date();
    const yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    
    // 날짜 범위 설정 (어제 00시 ~ 오늘 현재)
    const fromDate = new Date(yesterday);
    fromDate.setHours(0, 0, 0, 0);
    
    // 날짜 형식 변환 (YYYY-MM-DD)
    const fromDateStr = Utilities.formatDate(fromDate, "GMT", "yyyy-MM-dd");
    
    Logger.log(`뉴스 검색 날짜 범위: ${fromDateStr}`);
    
    // 키워드 인코딩
    const encodedKeyword = encodeURIComponent(keyword);
    
    // API URL 생성 - 수정된 부분: 단순화된 URL
    const apiUrl = `${CONFIG.NEWSAPI.ENDPOINT}everything?q=${encodedKeyword}&from=${fromDateStr}&sortBy=${CONFIG.NEWSAPI.SORT_BY}&apiKey=${apiKey}`;
    
    Logger.log(`NewsAPI 요청 URL: ${apiUrl}`);
    
    // 응답 로깅을 위한 변수
    let responseContent = "";
    
    // API 요청
    const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
    
    // 응답 상태 코드 확인
    const responseCode = response.getResponseCode();
    responseContent = response.getContentText();
    
    if (responseCode !== 200) {
      Logger.log(`NewsAPI 요청 실패 - 상태 코드: ${responseCode}, 응답: ${responseContent.substring(0, 200)}...`);
      return [];
    }
    
    // 응답 전체 내용 로깅 (디버깅을 위함)
    Logger.log(`NewsAPI 응답 내용: ${responseContent}`);
    
    // JSON 응답 파싱
    const jsonResponse = JSON.parse(responseContent);
    
    // API 상태 확인
    if (jsonResponse.status !== "ok") {
      Logger.log(`NewsAPI 응답 상태 오류: ${jsonResponse.status}, 메시지: ${jsonResponse.message || "알 수 없는 오류"}`);
      return [];
    }
    
    // 뉴스 아이템이 없는 경우
    if (!jsonResponse.articles || jsonResponse.articles.length === 0) {
      Logger.log(`NewsAPI 검색 결과 없음: ${keyword}`);
      return [];
    }
    
    Logger.log(`NewsAPI 검색 결과: ${jsonResponse.articles.length}개 기사 찾음`);
    
    // 응답 변환
    return parseNewsAPIResults(jsonResponse.articles, language);
    
  } catch (error) {
    Logger.log(`NewsAPI 검색 중 오류 발생: ${error.message}`);
    Logger.log(`스택: ${error.stack}`);
    return [];
  }
}

/**
 * NewsAPI 결과를 내부 형식으로 변환
 * @param {Array} articles - NewsAPI 기사 배열
 * @param {string} language - 언어 코드
 * @return {Array} 변환된 뉴스 아이템 배열
 */
function parseNewsAPIResults(articles, language) {
  const newsItems = [];
  
  for (const article of articles) {
    try {
      // 필수 필드 확인
      if (!article.title || !article.url) {
        continue;
      }
      
      // 발행일 파싱
      let pubDate = new Date();
      if (article.publishedAt) {
        pubDate = new Date(article.publishedAt);
      }
      
      // 소스 정보 추출
      let source = "NewsAPI";
      let sourceUrl = "";
      if (article.source) {
        source = article.source.name || "NewsAPI";
      }
      
      // 언어 플래그 설정
      const isEnglish = language === "en";
      
      // 뉴스 아이템 추가
      newsItems.push({
        title: article.title,
        link: article.url,
        description: article.description || article.content || "",
        pubDate: pubDate,
        pubDateStr: Utilities.formatDate(pubDate, "Asia/Seoul", "yyyyMMdd"),
        source: source,
        sourceUrl: sourceUrl,
        isGlobal: isEnglish,
        imageUrl: article.urlToImage || null,
        author: article.author || ""
      });
    } catch (itemError) {
      Logger.log(`뉴스 아이템 파싱 오류: ${itemError.message}`);
    }
  }
  
  return newsItems;
}

// ---------------------- Gemini AI 분석 함수 ----------------------

/**
 * Gemini AI 뉴스 필터링 함수 (배치 처리)
 * @param {Array} newsItems - 뉴스 아이템 배열
 * @param {string} topic - 주제
 * @param {string} relatedConcepts - 관련 개념
 * @param {string} apiKey - API 키
 * @param {string} newsType - 뉴스 유형
 * @return {Array} 필터링된 뉴스 아이템 배열
 */
function filterNewsByGeminiBatched(newsItems, topic, relatedConcepts, apiKey, newsType = "뉴스") {
  if (newsItems.length === 0) return [];
  
  try {
    // 배치 분할
    const batches = createBatches(newsItems, CONFIG.BATCH_SIZE, CONFIG.MAX_BATCHES);
    Logger.log(`총 ${newsItems.length}개 ${newsType}를 ${batches.length}개 배치로 분석합니다.`);
    
    // 각 배치별로 Gemini 분석 실행
    const allAnalyzedNews = [];
    
    for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
      const batch = batches[batchIndex];
      Logger.log(`배치 ${batchIndex + 1}/${batches.length} 분석 중 (${batch.length}개 ${newsType})`);
      
      // 배치 분석 (재시도 로직 포함)
      const analyzedBatch = analyzeNewsBatchWithRetries(
        batch, 
        topic, 
        relatedConcepts, 
        apiKey, 
        batchIndex * CONFIG.BATCH_SIZE, 
        newsType
      );
      
      // 분석된 배치 추가
      allAnalyzedNews.push(...analyzedBatch);
      
      // 배치 사이에 짧은 지연 시간 추가
      if (batchIndex < batches.length - 1) {
        Utilities.sleep(1500);
      }
    }
    
    // 후처리 (정렬, 중복 제거, 필터링)
    return postProcessAnalyzedNews(allAnalyzedNews, newsType);
    
  } catch (error) {
    Logger.log(`Gemini AI ${newsType} 배치 분석 중 오류 발생: ${error.message}`);
    Logger.log(`스택: ${error.stack}`);
    
    // 오류 발생 시 수동 필터링으로 대체
    return fallbackFilterNews(newsItems, topic, newsType);
  }
}

/**
 * 배치 생성 함수
 * @param {Array} items - 아이템 배열
 * @param {number} batchSize - 배치 크기
 * @param {number} maxBatches - 최대 배치 수
 * @return {Array} 배치 배열
 */
function createBatches(items, batchSize, maxBatches) {
  const batches = [];
  for (let i = 0; i < items.length; i += batchSize) {
    const end = Math.min(i + batchSize, items.length);
    batches.push(items.slice(i, end));
    
    // 최대 배치 수 제한
    if (batches.length >= maxBatches) break;
  }
  return batches;
}

/**
 * 뉴스 배치 분석 (재시도 로직 포함)
 * @param {Array} batch - 뉴스 배치
 * @param {string} topic - 주제
 * @param {string} relatedConcepts - 관련 개념
 * @param {string} apiKey - API 키
 * @param {number} startIndex - 시작 인덱스
 * @param {string} newsType - 뉴스 유형
 * @return {Array} 분석된 뉴스 배열
 */
function analyzeNewsBatchWithRetries(batch, topic, relatedConcepts, apiKey, startIndex, newsType) {
  let retryCount = 0;
  const maxRetries = CONFIG.MAX_RETRIES;
  let analyzedBatch = [];
  let success = false;
  
  while (retryCount <= maxRetries && !success) {
    try {
      // 분석 시도
      analyzedBatch = analyzeNewsBatchWithGemini(batch, topic, relatedConcepts, apiKey, startIndex, newsType);
      
      // 검증: 모든 항목에 aiRelevanceScore가 있는지 확인
      const allScored = analyzedBatch.every(item => typeof item.aiRelevanceScore === 'number');
      
      if (allScored) {
        success = true;
      } else {
        throw new Error("일부 뉴스 항목에 점수가 할당되지 않았습니다.");
      }
    } catch (error) {
      retryCount++;
      Logger.log(`배치 ${Math.floor(startIndex / CONFIG.BATCH_SIZE) + 1} 분석 중 오류, 재시도 ${retryCount}/${maxRetries}: ${error.message}`);
      
      if (retryCount <= maxRetries) {
        // 지수 백오프 (재시도마다 대기 시간 증가)
        Utilities.sleep(Math.pow(2, retryCount) * 1000);
      } else {
        // 최대 재시도 횟수 초과 시 기본 점수 부여
        Logger.log(`배치 ${Math.floor(startIndex / CONFIG.BATCH_SIZE) + 1} 분석 중 최대 재시도 횟수 초과`);
        // 임시 분석 결과 생성
        analyzedBatch = batch.map((item, index) => ({
          ...item,
          aiRelevanceScore: 7.5, // 임계값을 넘는 기본 점수 설정
          groupId: `failover_batch${startIndex}_${index}`,
          relevanceReason: "API 오류로 인한 기본 평가",
          newsType: "뉴스"
        }));
      }
    }
  }
  
  return analyzedBatch;
}

/**
 * 분석된 뉴스 후처리 (정렬, 중복 제거, 필터링) - 개선된 중복 제거 로직
 * @param {Array} allAnalyzedNews - 분석된 뉴스 배열
 * @param {string} newsType - 뉴스 유형
 * @return {Array} 후처리된 뉴스 배열
 */
function postProcessAnalyzedNews(allAnalyzedNews, newsType) {
  // 관련성 점수로 정렬
  allAnalyzedNews.sort((a, b) => b.aiRelevanceScore - a.aiRelevanceScore);
  
  // 중복 제거 (내용 기반 개선된 로직)
  const uniqueNews = removeDuplicateNews(allAnalyzedNews);
  
  // 최소 관련성 임계값 설정 (뉴스 유형에 따라 다른 임계값 적용)
  const threshold = newsType === "해외 뉴스" 
    ? CONFIG.RELEVANCE_THRESHOLD.GLOBAL 
    : CONFIG.RELEVANCE_THRESHOLD.DOMESTIC;
  
  const relevantNews = filterNewsByRelevance(uniqueNews, threshold, newsType);

  // 최대 반환 뉴스 수 제한
  const limitedResults = relevantNews.slice(0, CONFIG.MAX_NEWS_RESULTS);
  
  if (relevantNews.length > CONFIG.MAX_NEWS_RESULTS) {
    Logger.log(`최대 표시 개수 제한으로 ${relevantNews.length}개 중 상위 ${CONFIG.MAX_NEWS_RESULTS}개만 선택`);
  }
  
  // 뉴스가 없으면 임계값을 낮춰서 몇 개라도 반환 (fallback 메커니즘)
  if (limitedResults.length === 0 && uniqueNews.length > 0) {
    Logger.log(`${newsType}에서 임계값 ${threshold}점 이상 항목이 없어 임계값을 낮춰서 시도합니다.`);
    
    // 낮은 임계값 적용 (항상 몇 개라도 뉴스 표시)
    const lowerThreshold = newsType === "해외 뉴스" 
      ? CONFIG.RELEVANCE_THRESHOLD.FALLBACK_GLOBAL 
      : CONFIG.RELEVANCE_THRESHOLD.FALLBACK_DOMESTIC;
    
    const fallbackNews = filterNewsByRelevance(uniqueNews, lowerThreshold, newsType).slice(0, 3);
    
    Logger.log(`낮은 임계값 ${lowerThreshold}점 이상으로 ${fallbackNews.length}개 뉴스 선택`);
    return fallbackNews;
  }

  // 필터링된 뉴스 반환
  return limitedResults;
}

/**
 * 관련성 점수로 뉴스 필터링
 * @param {Array} newsItems - 뉴스 아이템 배열
 * @param {number} threshold - 임계값
 * @param {string} newsType - 뉴스 유형
 * @return {Array} 필터링된 뉴스 아이템 배열
 */
function filterNewsByRelevance(newsItems, threshold, newsType) {
  return newsItems.filter(item => {
    // 점수가 문자열이거나 숫자인 경우 모두 처리
    const score = typeof item.aiRelevanceScore === 'string' 
                ? parseFloat(item.aiRelevanceScore) 
                : (Number(item.aiRelevanceScore) || 0);
    
    const passes = score >= threshold;
    
    // 디버깅용 로깅 추가 (소수점 한 자리까지 표시)
    if (passes) {
      Logger.log(`${newsType} 통과: "${item.title.substring(0, 30)}..." - 점수: ${score.toFixed(1)}, 임계값: ${threshold.toFixed(1)}`);
    } else {
      Logger.log(`${newsType} 미통과: "${item.title.substring(0, 30)}..." - 점수: ${score.toFixed(1)}, 임계값: ${threshold.toFixed(1)}`);
    }
    
    return passes;
  });
}

/**
 * 단일 배치 뉴스 분석 함수 (관련 개념 활용)
 * @param {Array} newsBatch - 뉴스 배치
 * @param {string} topic - 주제
 * @param {string} relatedConcepts - 관련 개념
 * @param {string} apiKey - API 키
 * @param {number} startIndex - 시작 인덱스
 * @param {string} newsType - 뉴스 유형
 * @return {Array} 분석된 뉴스 배열
 */
function analyzeNewsBatchWithGemini(newsBatch, topic, relatedConcepts, apiKey, startIndex, newsType = "뉴스") {
  try {
    const modelName = getSelectedGeminiModel();
    const prompt = createAnalysisPrompt(topic, relatedConcepts, newsType, newsBatch, startIndex);
    
    Logger.log(`Gemini 분석 요청 - 주제: ${topic}, 뉴스 개수: ${newsBatch.length}, 모델: ${modelName}`);
    
    // Gemini API 호출
    const response = callGeminiAPI(prompt, apiKey, modelName);
    
    // 응답에서 JSON 부분 추출
    const jsonMatch = extractJsonFromText(response);
    
    Logger.log(`Gemini 응답 JSON 추출 성공 여부: ${jsonMatch ? "성공" : "실패"}`);
    
    if (!jsonMatch) {
      Logger.log("배치 분석: Gemini API 응답에서 JSON을 추출할 수 없습니다.");
      Logger.log(`응답 내용 일부: ${response.substring(0, 500)}...`);
      return createDefaultAnalysisResults(newsBatch, startIndex);
    }
    
    // JSON 파싱 및 결과 처리 전
    Logger.log(`Gemini 응답 원본 JSON: ${jsonMatch}`);
    const parsedResults = parseAnalysisResults(jsonMatch, newsBatch, startIndex);
    Logger.log(`파싱 결과 (첫 번째 항목): ${JSON.stringify(parsedResults[0])}`);
    return parsedResults;
    
  } catch (error) {
    Logger.log(`단일 배치 분석 중 오류 발생: ${error.message}`);
    Logger.log(`스택: ${error.stack}`);
    
    // 오류 발생 시 기본 점수와 고유 그룹 ID 할당
    return createDefaultAnalysisResults(newsBatch, startIndex);
  }
}

/**
 * 뉴스 분석용 프롬프트 생성
 * @param {string} topic - 주제
 * @param {string} relatedConcepts - 관련 개념
 * @param {string} newsType - 뉴스 유형
 * @param {Array} newsBatch - 뉴스 배치
 * @param {number} startIndex - 시작 인덱스
 * @return {string} 분석 프롬프트
 */
function createAnalysisPrompt(topic, relatedConcepts, newsType, newsBatch, startIndex) {
  let prompt = `다음은 "${topic}" 주제와 관련된 ${newsType} 기사 목록입니다. 각 기사가 이 주제와 실질적으로 얼마나 관련이 있는지 의미적 관련성을 평가해주세요.

  특별히 다음 사항에 주의해주세요:
  1. 정확히 같은 단어가 사용되지 않더라도, 의미적으로 관련된 내용이면 높은 관련성 점수를 부여하세요.
  2. 기사가 "${topic}"의 하위 주제나 관련 기술을 다루고 있으면 관련성이 높습니다.
  3. 단순히 키워드가 언급되었다고 관련성이 높은 것은 아닙니다. 실제 내용이 주제와 얼마나 관련되어 있는지 평가해주세요.
  4. 주식과 관련한 뉴스는 관련성을 낮게 평가하세요.`;

  // 뉴스 유형별 추가 지침
  if (newsType === "해외 뉴스") {
    prompt += `
      4. 이 기사들은 해외 뉴스입니다. 글로벌 시장 동향이나 해외 기업 활동도 중요한 정보로 평가해주세요.
      5. 직접적인 언급이 없더라도 글로벌 맥락에서 "${topic}"와 연관성이 있으면 관련성 점수를 높게 주세요.
      6. 해외 뉴스는 번역이나 문화적 차이로 인해 문맥이 명확하지 않을 수 있으니, 관련성이 의심스러울 때는 7점 이상을 부여해주세요.
      7. 주식과 관련한 뉴스는 관련성을 낮게 평가하세요.`;
  } else {
    prompt += `
      4. 기사가 국내 "${topic}" 업계, 정책, 기업 활동 등을 구체적으로 다루는 경우 관련성이 높습니다.
      5. 관련성 점수 7점 이상은 해당 주제와 직접적으로 관련이 있고 중요한 내용을 담고 있는 기사에만 부여해주세요.
      6. 10점은 해당 주제의 핵심적인 내용을 다루는 매우 중요한 기사에만 부여하세요.
      7. 4. 주식과 관련한 뉴스는 관련성을 낮게 평가하세요.`;
  }

  // 관련 개념이 있으면 프롬프트에 추가
  if (relatedConcepts && relatedConcepts.trim() !== "") {
    const conceptsList = relatedConcepts.split(",").map(c => c.trim()).filter(c => c !== "");
    if (conceptsList.length > 0) {
      prompt += `\n\n이 주제와 관련된 개념들은 다음과 같습니다: ${conceptsList.join(", ")}`;
      prompt += `\n위의 관련 개념들이 뉴스에 언급되거나 관련된 내용이 있으면 관련성이 높다고 볼 수 있습니다.`;
    }
  }
  
  // JSON 응답 형식 명확화
  prompt += `\n\n각 기사에 대해 1-10 척도로 관련성 점수를 부여하고, 유사한 내용을 다루는 기사는 같은 그룹으로 분류해주세요.

    다음과 같은 JSON 배열 형식으로만 응답해주세요:
    [
      {"id": 0, "relevanceScore": 8.5, "groupId": "A", "relevanceReason": "간략한 이유", "newsType": "연구"},
      {"id": 1, "relevanceScore": 5.2, "groupId": "B", "relevanceReason": "간략한 이유", "newsType": "정책"}
    ]
    
    응답은 반드시 이 형식의 JSON만 포함해야 합니다. 다른 텍스트는 포함하지 마세요.
    
    newsType 속성값은 다음 중 하나를 선택하세요: "연구", "정책", "투자", "제품", "시장", "기술".
    
    `;

  // 뉴스 항목 추가
  prompt += `\n\n분석할 뉴스 기사:`;
  newsBatch.forEach((item, index) => {
    const globalIndex = startIndex + index;
    prompt += `\n[${globalIndex}] 제목: "${item.title}"\n`;
    prompt += `내용: "${item.description || '내용 없음'}"\n`;
  });
  
  return prompt;
}

/**
 * 기본 분석 결과 생성
 * @param {Array} newsBatch - 뉴스 배치
 * @param {number} startIndex - 시작 인덱스
 * @return {Array} 기본 분석 결과
 */
function createDefaultAnalysisResults(newsBatch, startIndex) {
  return newsBatch.map((item, index) => {
    // 랜덤한 관련성 점수 생성 (7.0-9.0 사이)
    const score = 7.0 + Math.random() * 2.0;
    
    return {
      ...item,
      aiRelevanceScore: score,
      groupId: `unique_batch${startIndex}_${index}`,
      relevanceReason: "분석 결과 없음",
      newsType: ["연구", "정책", "시장", "제품", "투자", "기술"][Math.floor(Math.random() * 6)]
    };
  });
}

/**
 * 분석 결과 파싱
 * @param {string} jsonMatch - JSON 문자열
 * @param {Array} newsBatch - 뉴스 배치
 * @param {number} startIndex - 시작 인덱스
 * @return {Array} 파싱된 분석 결과
 */
function parseAnalysisResults(jsonMatch, newsBatch, startIndex) {
  try {
    // 깨끗한 JSON 문자열을 파싱
    const analysisResult = JSON.parse(jsonMatch);
    
    Logger.log(`Gemini 응답 JSON 분석 결과 - 항목 수: ${analysisResult.length}`);
    
    if (!Array.isArray(analysisResult)) {
      Logger.log("파싱된 JSON이 배열이 아닙니다. 기본값 사용.");
      return createDefaultAnalysisResults(newsBatch, startIndex);
    }
    
    // 디버깅: 첫 번째 항목 로깅
    if (analysisResult.length > 0) {
      Logger.log(`첫 번째 분석 결과 항목: ${JSON.stringify(analysisResult[0])}`);
    }
    
    // 분석 결과를 뉴스 아이템에 추가
    return newsBatch.map((item, index) => {
      const globalIndex = startIndex + index;
      
      // 해당 globalIndex의 분석 결과 찾기
      const analysis = analysisResult.find(a => {
        // id가 문자열 또는 숫자일 수 있으므로 두 경우 모두 처리
        const analysisId = typeof a.id === 'string' ? parseInt(a.id, 10) : a.id;
        return analysisId === globalIndex;
      });
      
      if (!analysis) {
        Logger.log(`ID ${globalIndex}에 대한 분석 결과를 찾을 수 없습니다.`);
        return {
          ...item,
          aiRelevanceScore: 7.5, // 기본값
          groupId: `unique_batch${startIndex}_${index}`,
          relevanceReason: "분석 결과 없음",
          newsType: "뉴스"
        };
      }
      
      // 정확한 관련성 점수 및 타입 추출
      const score = parseRelevanceScore(analysis.relevanceScore);
      const newsType = analysis.newsType || "뉴스";
      
      Logger.log(`뉴스 아이템 분석 결과 - ID: ${globalIndex}, 점수: ${score}, 타입: ${newsType}`);
      
      return {
        ...item,
        aiRelevanceScore: score,
        groupId: analysis.groupId || `unique_batch${startIndex}_${index}`,
        relevanceReason: analysis.relevanceReason || "",
        newsType: newsType
      };
    });
  } catch (parseError) {
    Logger.log(`JSON 파싱 오류: ${parseError.message}, 원본 JSON: ${jsonMatch.substring(0, 200)}...`);
    return createDefaultAnalysisResults(newsBatch, startIndex);
  }
}

/**
 * 관련성 점수 파싱
 * @param {string|number} scoreValue - 점수 값
 * @return {number} 파싱된 점수
 */
function parseRelevanceScore(scoreValue) {
  try {
    const score = typeof scoreValue === 'string' 
              ? parseFloat(scoreValue) 
              : (Number(scoreValue) || 7.5);
              
    // NaN이나 범위를 벗어나는 값 처리
    if (isNaN(score) || score < 1 || score > 10) {
      Logger.log(`유효하지 않은 점수 (${scoreValue}), 기본값 7.5 사용`);
      return 7.5;
    }
    
    return score;
  } catch (e) {
    Logger.log(`점수 변환 오류: ${e.message}, 기본값 7.5 사용`);
    return 7.5;
  }
}

/**
 * Gemini API 호출 함수 (재시도 로직 포함)
 * @param {string} prompt - 프롬프트
 * @param {string} apiKey - API 키
 * @param {string} modelName - 모델 이름
 * @return {string} API 응답
 */
function callGeminiAPI(prompt, apiKey, modelName = null) {
  // 모델이 지정되지 않은 경우 기본값 사용
  if (!modelName) {
    modelName = getSelectedGeminiModel();
  }
  
  const apiEndpoint = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent`;
  
  let retryCount = 0;
  const maxRetries = CONFIG.MAX_RETRIES;
  
  while (retryCount < maxRetries) {
    try {
      // 프롬프트 길이 제한 (토큰 제한 방지)
      const truncatedPrompt = prompt.length > 20000 ? prompt.substring(0, 20000) : prompt;
      
      Logger.log(`Gemini API 요청 - 모델: ${modelName}, 프롬프트 길이: ${truncatedPrompt.length}`);
      
      const response = UrlFetchApp.fetch(`${apiEndpoint}?key=${apiKey}`, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{
            parts: [{
              text: truncatedPrompt
            }]
          }],
          generationConfig: {
            temperature: 0.1,
            topP: 0.8,
            topK: 40,
            maxOutputTokens: 2048,
            stopSequences: []
          },
          safetySettings: [
            {
              category: "HARM_CATEGORY_DANGEROUS_CONTENT",
              threshold: "BLOCK_NONE"
            },
            {
              category: "HARM_CATEGORY_HATE_SPEECH",
              threshold: "BLOCK_NONE"
            },
            {
              category: "HARM_CATEGORY_HARASSMENT",
              threshold: "BLOCK_NONE"
            },
            {
              category: "HARM_CATEGORY_SEXUALLY_EXPLICIT",
              threshold: "BLOCK_NONE"
            }
          ]
        }),
        muteHttpExceptions: true
      });
      
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      // 응답 내용 로깅 (처음 500자 정도만)
      Logger.log(`Gemini API 응답 코드: ${responseCode}`);
      Logger.log(`Gemini API 응답 내용 (처음 500자): ${responseText.substring(0, 500)}...`);
      
      // 성공 시 응답 반환
      if (responseCode === 200) {
        const responseData = JSON.parse(responseText);
        if (responseData.candidates && responseData.candidates.length > 0 && 
            responseData.candidates[0].content && responseData.candidates[0].content.parts) {
          const textResponse = responseData.candidates[0].content.parts[0].text;
          Logger.log(`Gemini 응답 텍스트 (처음 300자): ${textResponse.substring(0, 300)}...`);
          return textResponse;
        } else {
          Logger.log("Gemini API 응답 형식이 예상과 다릅니다.");
          Logger.log(`응답 데이터: ${JSON.stringify(responseData).substring(0, 500)}...`);
          return "";
        }
      }
      
      // API 오류 코드에 따른 처리
      if (responseCode === 503 || responseCode === 429) {
        // 서비스 불가(503) 또는 할당량 초과(429)
        retryCount++;
        Logger.log(`Gemini API 호출 실패 (${responseCode}): 재시도 ${retryCount}/${maxRetries}`);
        
        if (retryCount < maxRetries) {
          // 지수 백오프 (재시도마다 대기 시간 증가)
          const waitTime = Math.pow(2, retryCount) * 1000; // 2초, 4초, 8초...
          Utilities.sleep(waitTime);
        }
      } else {
        // 다른 오류는 로그 기록 후 빈 문자열 반환
        Logger.log(`Gemini API 오류 (${responseCode}): ${responseText}`);
        return "";
      }
    } catch (error) {
      retryCount++;
      Logger.log(`Gemini API 호출 중 예외 발생: ${error.message}, 재시도 ${retryCount}/${maxRetries}`);
      
      if (retryCount < maxRetries) {
        // 예외 발생 시도 재시도
        const waitTime = Math.pow(2, retryCount) * 1000;
        Utilities.sleep(waitTime);
      } else {
        // 최대 재시도 횟수 초과하면 빈 문자열 반환
        return "";
      }
    }
  }
  
  // 모든 재시도 실패 시 빈 문자열 반환
  return "";
}

/**
 * JSON 추출 헬퍼 함수
 * @param {string} text - 텍스트
 * @return {string|null} 추출된 JSON 문자열
 */
function extractJsonFromText(text) {
  try {
    // 백틱과 json 태그 제거
    let cleanText = text.replace(/```json\n|\n```/g, '');
    
    // 배열 매칭 (가장 일반적인 Gemini 응답 형식)
    if (cleanText.trim().startsWith('[') && cleanText.trim().endsWith(']')) {
      return cleanText.trim();
    }
    
    // 기존 로직 (객체 매칭)
    let openBraces = 0;
    let startIdx = -1;
    let endIdx = -1;
    
    for (let i = 0; i < cleanText.length; i++) {
      if (cleanText[i] === '{') {
        if (openBraces === 0) {
          startIdx = i;
        }
        openBraces++;
      } else if (cleanText[i] === '}') {
        openBraces--;
        if (openBraces === 0) {
          endIdx = i;
          break;
        }
      }
    }
    
    if (startIdx !== -1 && endIdx !== -1) {
      return cleanText.substring(startIdx, endIdx + 1);
    }
    
    // 정규식 백업 방법
    const jsonMatch = cleanText.match(/(\{[\s\S]*\})|(\[[\s\S]*\])/);
    if (jsonMatch) {
      return jsonMatch[0];
    }
    
    return null;
  } catch (e) {
    Logger.log(`JSON 추출 중 오류: ${e.message}`);
    return null;
  }
}

/**
 * 영어 뉴스 제목과 요약 번역 함수
 * @param {Array} newsItems - 뉴스 아이템 배열
 * @param {string} apiKey - API 키
 * @return {Array} 번역된 뉴스 아이템 배열
 */
function translateNewsWithGemini(newsItems, apiKey) {
  if (newsItems.length === 0) return [];
  
  try {
    // 최대 처리할 뉴스 수 제한
    const maxNewsToTranslate = Math.min(newsItems.length, 5);
    const newsToTranslate = newsItems.slice(0, maxNewsToTranslate);
    
    Logger.log(`영어 뉴스 ${newsToTranslate.length}개 번역 시작`);
    
    // 각 뉴스 아이템에 대해 번역 수행
    const translatedNewsItems = [];
    
    for (let i = 0; i < newsToTranslate.length; i++) {
      const news = newsToTranslate[i];
      const translationResult = translateSingleNewsItem(news, apiKey);
      translatedNewsItems.push(translationResult);
      
      // API 요청 사이에 짧은 지연 추가
      if (i < newsToTranslate.length - 1) {
        Utilities.sleep(1000);
      }
    }
    
    Logger.log(`영어 뉴스 번역 완료: ${translatedNewsItems.length}개`);
    return translatedNewsItems;
    
  } catch (error) {
    Logger.log(`뉴스 번역 중 오류 발생: ${error.message}`);
    return newsItems; // 오류 발생 시 원본 반환
  }
}

/**
 * 단일 뉴스 아이템 번역
 * @param {Object} news - 뉴스 아이템
 * @param {string} apiKey - API 키
 * @return {Object} 번역된 뉴스 아이템
 */
function translateSingleNewsItem(news, apiKey) {
  // 번역 및 요약을 위한 프롬프트 생성
  const prompt = `다음 영어 뉴스를 한국어로 번역하고 2~3문장으로 요약해주세요.
  
  원문 제목: "${news.title}"
  원문 내용: "${news.description}"
  
  다음 JSON 형식으로 응답해주세요:
  {
    "translatedTitle": "한국어로 번역된 제목",
    "summary": "한국어로 2~3문장 요약된 내용"
  }
  
  JSON 형식만 반환해주세요.`;
  
  // Gemini API 호출
  const response = callGeminiAPI(prompt, apiKey);
  
  // 응답 파싱
  try {
    // JSON 형식 추출
    const jsonMatch = extractJsonFromText(response);
    if (!jsonMatch) {
      Logger.log(`뉴스 번역 중 JSON 형식을 찾을 수 없습니다: ${response.substring(0, 200)}...`);
      // 번역 실패 시 원본 그대로 반환
      return news;
    }
    
    const translationData = JSON.parse(jsonMatch);
    
    Logger.log(`번역 성공: "${news.title.substring(0, 30)}..." => "${translationData.translatedTitle?.substring(0, 30)}..."`);
    
    // 번역된 정보로 뉴스 항목 업데이트
    return {
      ...news,
      title: translationData.translatedTitle || news.title,
      aiSummary: translationData.summary || news.description
    };
    
  } catch (parseError) {
    Logger.log(`번역 결과 파싱 오류: ${parseError.message}`);
    // 번역 실패 시 원본 그대로 반환
    return news;
  }
}

/**
 * API 오류 시 대체 필터링 함수
 * @param {Array} newsItems - 뉴스 아이템 배열
 * @param {string} topic - 주제
 * @param {string} newsType - 뉴스 유형
 * @return {Array} 필터링된 뉴스 아이템 배열
 */
function fallbackFilterNews(newsItems, topic, newsType) {
  Logger.log(`API 오류로 인해 수동 필터링으로 전환 (${newsType})`);
  
  // 키워드 기반 필터링
  const keywords = topic.split(/\s+/).filter(word => word.length > 1);
  
  // 점수 계산 및 정렬
  const scoredNews = newsItems.map(item => {
    const score = calculateKeywordScore(item.title, item.description, keywords);
    return {
      ...item,
      aiRelevanceScore: score,
      relevanceReason: "키워드 기반 수동 평가",
      newsType: "뉴스"
    };
  }).filter(item => item.aiRelevanceScore > 0);
  
  // 점수 정규화 (1-10 스케일로)
  const maxScore = Math.max(...scoredNews.map(item => item.aiRelevanceScore), 1);
  scoredNews.forEach(item => {
    item.aiRelevanceScore = Math.min(10, Math.max(7.0, Math.round(item.aiRelevanceScore / maxScore * 10)));
  });
  
  // 정렬 및 최대 항목 제한
  scoredNews.sort((a, b) => b.aiRelevanceScore - a.aiRelevanceScore);
  
  // 최대 5개로 제한
  const limitCount = Math.min(5, scoredNews.length);
  const result = scoredNews.slice(0, limitCount);
  
  Logger.log(`수동 필터링 결과: ${scoredNews.length}개 중 상위 ${result.length}개 선택`);
  
  return result;
}

/**
 * 키워드 기반 관련성 점수 계산
 * @param {string} title - 제목
 * @param {string} description - 설명
 * @param {Array} keywords - 키워드 배열
 * @return {number} 관련성 점수
 */
function calculateKeywordScore(title, description, keywords) {
  let score = 0;
  const text = (title + ' ' + (description || '')).toLowerCase();
  
  // 각 키워드가 존재하는지 확인
  keywords.forEach(keyword => {
    if (text.includes(keyword.toLowerCase())) {
      score += 2;
    }
  });
  
  // 제목에 키워드가 있으면 가산점
  keywords.forEach(keyword => {
    if (title.toLowerCase().includes(keyword.toLowerCase())) {
      score += 1;
    }
  });
  
  return score;
}

/**
 * 선택된 Gemini 모델 가져오기
 * @return {string} 모델 이름
 */
function getSelectedGeminiModel() {
  return PropertiesService.getUserProperties().getProperty(CONFIG.GEMINI_MODEL_PROPERTY) || CONFIG.DEFAULT_GEMINI_MODEL;
}

// ---------------------- UI 인터페이스 함수 ----------------------

/**
 * 초기 스프레드시트 설정 함수 (네이버 API 열 추가)
 */
function setupSpreadsheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 헤더 설정
  const headers = [
    ["A1", "주제(키워드)"],
    ["B1", "이메일"],
    ["C1", "네이버 Client ID"],
    ["D1", "네이버 Client Secret"],
    ["E1", "Google AI API 키"],
    ["F1", "관련 개념(쉼표로 구분)"],
    ["G1", "해외 뉴스 포함(Y/N)"],
    ["H1", "해외 뉴스 언어(en)"],
    ["I1", "해외 뉴스 검색 키워드(영어)"],
    ["J1", "NewsAPI 키"]
  ];
  
  headers.forEach(([cell, value]) => {
    sheet.getRange(cell).setValue(value);
  });
  
  // 서식 지정
  sheet.getRange("A1:J1").setFontWeight("bold");
  sheet.getRange("A:J").setVerticalAlignment("middle");
  
  // 열 너비 설정
  const columnWidths = [
    [1, 150], [2, 200], [3, 180], [4, 180], [5, 200],
    [6, 250], [7, 150], [8, 150], [9, 200], [10, 200]
  ];
  
  columnWidths.forEach(([column, width]) => {
    sheet.setColumnWidth(column, width);
  });
  
  // 예시 데이터
  sheet.getRange("A2").setValue("배터리 소재");
  sheet.getRange("B2").setValue("user1@example.com");
  // C2와 D2에 네이버 API 정보 입력 예정
  sheet.getRange("F2").setValue("양극재,음극재,전해질,분리막,황산니켈,리튬,코발트");
  sheet.getRange("G2").setValue("Y");
  sheet.getRange("H2").setValue("en");
  sheet.getRange("I2").setValue("battery materials,lithium,cathode,anode");
  // J2에는 NewsAPI 키 입력 예정
  
  // 두 번째 예시 행
  sheet.getRange("A3").setValue("EV");
  sheet.getRange("B3").setValue("user2@example.com");
  sheet.getRange("F3").setValue("전기차,충전,자율주행,테슬라,현대차,기아");
  sheet.getRange("G3").setValue("Y");
  sheet.getRange("I3").setValue("electric vehicle,EV,charging");
  
  Logger.log("스프레드시트 초기 설정이 완료되었습니다.");
}

/**
 * 트리거 설정을 위한 함수
 */
function createDailyTrigger() {
  // 기존 트리거 삭제
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "sendNewsletterEmail") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // 매일 오전 8시에 실행되는 트리거 생성
  ScriptApp.newTrigger("sendNewsletterEmail")
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
    
  Logger.log("매일 오전 8시에 뉴스레터를 보내는 트리거가 설정되었습니다.");
}

/**
 * 테스트용 함수 (단일 주제만 처리하여 이메일 발송)
 */
function testNewsletterForTopic() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selectedCell = sheet.getActiveCell();
  const row = selectedCell.getRow();
  
  // 첫 번째 행이거나 빈 행인 경우 처리하지 않음
  if (row <= 1 || !sheet.getRange(`A${row}`).getValue()) {
    Browser.msgBox("알림", "주제가 있는 행을 선택한 후 실행해주세요.", Browser.Buttons.OK);
    return;
  }
  
  // 네이버 API 정보 가져오기
  const naverClientId = sheet.getRange("C2").getValue();
  const naverClientSecret = sheet.getRange("D2").getValue();
  
  // 네이버 API 정보 확인
  if (!naverClientId || !naverClientSecret) {
    Browser.msgBox("오류", "네이버 API 인증 정보가 없습니다. C2와 D2 셀을 확인해주세요.", Browser.Buttons.OK);
    return;
  }
  
  // 선택된 주제 정보 가져오기
  const topic = sheet.getRange(`A${row}`).getValue();
  const email = sheet.getRange(`B${row}`).getValue() || CONFIG.DEFAULT_EMAIL;
  const relatedConcepts = sheet.getRange(`F${row}`).getValue() || "";
  const includeGlobalNews = (sheet.getRange(`G${row}`).getValue() || "N").toUpperCase() === "Y";
  const englishKeyword = sheet.getRange(`I${row}`).getValue() || "";
  const newsApiKey = sheet.getRange(`J${row}`).getValue() || "";
  
  // API 인증 정보 가져오기
  const googleApiKey = sheet.getRange(`E${row}`).getValue() || sheet.getRange("E2").getValue();
  
  // 필수 정보 확인
  if (!topic) {
    Browser.msgBox("오류", "주제가 필요합니다.", Browser.Buttons.OK);
    return;
  }
  
  if (!googleApiKey) {
    Browser.msgBox("오류", "Google AI API 키가 필요합니다.", Browser.Buttons.OK);
    return;
  }
  
  if (includeGlobalNews && !newsApiKey) {
    Browser.msgBox("알림", "NewsAPI 키가 없어 해외 뉴스는 검색되지 않습니다.", Browser.Buttons.OK);
  }
  
  // 테스트용 단일 주제 뉴스레터 생성 및 발송
  createAndSendNewsletter(
    [{
      topic: topic,
      relatedConcepts: relatedConcepts,
      includeGlobalNews: includeGlobalNews,
      englishKeyword: englishKeyword,
      newsApiKey: newsApiKey
    }],
    [email],
    googleApiKey,
    naverClientId,
    naverClientSecret
  );
  
  Browser.msgBox("성공", `${topic} 주제에 대한 뉴스레터를 ${email}로 발송했습니다.`, Browser.Buttons.OK);
}

/**
 * 네이버 뉴스 API 테스트 함수
 */
function testNaverNewsSearch() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selectedCell = sheet.getActiveCell();
  const row = selectedCell.getRow();
  
  // 첫 번째 행이거나 빈 행인 경우 처리하지 않음
  if (row <= 1 || !sheet.getRange(`A${row}`).getValue()) {
    Browser.msgBox("알림", "주제가 있는 행을 선택한 후 실행해주세요.", Browser.Buttons.OK);
    return;
  }
  
  // 네이버 API 정보 가져오기
  const naverClientId = sheet.getRange("C2").getValue();
  const naverClientSecret = sheet.getRange("D2").getValue();
  
  // 네이버 API 정보 확인
  if (!naverClientId || !naverClientSecret) {
    Browser.msgBox("오류", "네이버 API 인증 정보가 없습니다. C2와 D2 셀을 확인해주세요.", Browser.Buttons.OK);
    return;
  }
  
  // 선택된 주제 정보 가져오기
  const topic = sheet.getRange(`A${row}`).getValue();
  
  try {
    // 네이버 뉴스 검색 테스트
    const naverNewsItems = searchNaverNews(topic, naverClientId, naverClientSecret);
    
    // 메시지 작성
    let message = `네이버 뉴스 검색 결과:\n\n`;
    message += `- 총 검색 결과: ${naverNewsItems.length}개\n`;
    
    // 첫 번째 뉴스 정보 표시
    if (naverNewsItems.length > 0) {
      const firstNews = naverNewsItems[0];
      message += `\n첫 번째 뉴스:\n`;
      message += `- 제목: ${firstNews.title}\n`;
      message += `- 출처: ${firstNews.source}\n`;
      message += `- 날짜: ${Utilities.formatDate(firstNews.pubDate, "Asia/Seoul", "yyyy-MM-dd HH:mm")}\n`;
    }
    
    Browser.msgBox("네이버 뉴스 검색 테스트", message, Browser.Buttons.OK);
    
  } catch (error) {
    Browser.msgBox("오류", `네이버 뉴스 검색 중 오류가 발생했습니다: ${error.message}`, Browser.Buttons.OK);
  }
}

/**
 * NewsAPI 테스트 함수
 */
function testNewsApiSearch() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selectedCell = sheet.getActiveCell();
  const row = selectedCell.getRow();
  
  // 첫 번째 행이거나 빈 행인 경우 처리하지 않음
  if (row <= 1 || !sheet.getRange(`A${row}`).getValue()) {
    Browser.msgBox("알림", "주제가 있는 행을 선택한 후 실행해주세요.", Browser.Buttons.OK);
    return;
  }
  
  // 선택된 주제 정보 가져오기
  const topic = sheet.getRange(`A${row}`).getValue();
  const englishKeyword = sheet.getRange(`I${row}`).getValue() || topic;
  const newsApiKey = sheet.getRange(`J${row}`).getValue();
  
  // NewsAPI 키 확인
  if (!newsApiKey) {
    Browser.msgBox("오류", "NewsAPI 키가 필요합니다. J열에 입력해주세요.", Browser.Buttons.OK);
    return;
  }
  
  // NewsAPI 검색 테스트
  try {
    const englishResults = searchNewsAPI(englishKeyword, "en", newsApiKey);
    
    // 메시지 작성
    let message = `NewsAPI 검색 결과:\n\n`;
    message += `- 영어 뉴스 (키워드: ${englishKeyword}): ${englishResults.length}개\n`;
    
    // 첫 번째 결과 표시
    if (englishResults.length > 0) {
      const firstNews = englishResults[0];
      message += `\n첫 번째 영어 뉴스:\n`;
      message += `- 제목: ${firstNews.title}\n`;
      message += `- 출처: ${firstNews.source}\n`;
      message += `- 날짜: ${Utilities.formatDate(firstNews.pubDate, "Asia/Seoul", "yyyy-MM-dd HH:mm")}\n`;
    }
    
    Browser.msgBox("NewsAPI 검색 테스트", message, Browser.Buttons.OK);
    
  } catch (error) {
    Browser.msgBox("오류", `NewsAPI 검색 중 오류가 발생했습니다: ${error.message}`, Browser.Buttons.OK);
  }
}

/**
 * 메뉴 추가 함수
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('뉴스레터')
    .addItem('초기 설정하기', 'setupSpreadsheet')
    .addItem('전체 뉴스레터 발송', 'sendNewsletterEmail')
    .addItem('선택한 주제만 테스트 발송', 'testNewsletterForTopic')
    .addSeparator()
    .addItem('매일 자동 발송 설정', 'createDailyTrigger')
    .addSeparator()
    .addSubMenu(ui.createMenu('도구')
      .addItem('네이버 뉴스 API 테스트', 'testNaverNewsSearch')
      .addItem('NewsAPI 검색 테스트', 'testNewsApiSearch')
    )
    .addToUi();
}

/**
 * 개선된 중복 뉴스 제거 함수
 * 제목의 유사도와 내용 유사도를 모두 고려하여 중복 뉴스를 필터링
 * @param {Array} newsItems - 뉴스 아이템 배열
 * @return {Array} 중복이 제거된 뉴스 아이템 배열
 */
function removeDuplicateNews(newsItems) {
  if (!newsItems || newsItems.length === 0) return [];
  
  // 결과 배열 초기화
  const uniqueNews = [];
  // 이미 처리된 뉴스의 핵심 내용 추적
  const processedContents = new Set();
  
  Logger.log(`중복 제거 전 뉴스 항목 수: ${newsItems.length}`);
  
  for (const item of newsItems) {
    // 제목과 내용에서 핵심 키워드 추출
    const titleSignature = getNormalizedSignature(item.title);
    const contentSignature = item.description ? getNormalizedSignature(item.description) : "";
    
    // 중복 검사를 위한 식별 키 생성 (기사 내용 포함)
    const duplicateCheckKey = generateDuplicateCheckKey(titleSignature, contentSignature);
    
    // 중복 검사
    if (!processedContents.has(duplicateCheckKey)) {
      // 중복이 아닌 경우 추가
      uniqueNews.push(item);
      processedContents.add(duplicateCheckKey);
      
      // 제목 기반 추가 중복 검사를 위한 키워드 추출
      const titleKeywords = extractKeywords(item.title);
      // 핵심 키워드 조합도 중복 체크에 추가
      for (const keyword of titleKeywords) {
        if (keyword.length > 5) { // 의미 있는 길이의 키워드만 추가
          processedContents.add(keyword);
        }
      }
    } else {
      Logger.log(`중복 뉴스 제거: "${item.title.substring(0, 30)}..."`);
    }
  }
  
  Logger.log(`중복 제거 후 뉴스 항목 수: ${uniqueNews.length}, 제거된 항목 수: ${newsItems.length - uniqueNews.length}`);
  
  return uniqueNews;
}

/**
 * 텍스트의 정규화된 식별자 생성
 * 공백 제거, 소문자 변환, 특수문자 제거 후 핵심 부분만 추출
 * @param {string} text - 원본 텍스트
 * @return {string} 정규화된 식별자
 */
function getNormalizedSignature(text) {
  if (!text) return "";
  
  // 괄호 내용 및 특수문자 제거, 소문자 변환, 공백 제거
  return text
    .replace(/\(.*?\)/g, "") // 괄호 내용 제거
    .replace(/[^\w\s가-힣]/g, "") // 한글과 영숫자가 아닌 문자 제거
    .replace(/\s+/g, "") // 공백 제거
    .toLowerCase(); // 소문자 변환
}

/**
 * 중복 검사용 키 생성
 * 제목과 내용의 특징을 조합하여 고유한 식별자 생성
 * @param {string} titleSignature - 정규화된 제목
 * @param {string} contentSignature - 정규화된 내용
 * @return {string} 중복 검사용 키
 */
function generateDuplicateCheckKey(titleSignature, contentSignature) {
  // 제목이 짧으면 전체 사용, 길면 앞부분만 사용
  const titlePart = titleSignature.substring(0, 50);
  
  // 내용이 있으면 내용의 일부를 조합
  let contentPart = "";
  if (contentSignature && contentSignature.length > 0) {
    contentPart = contentSignature.substring(0, 50);
  }
  
  // 제목만으로 체크할 경우
  if (!contentPart) {
    return titlePart;
  }
  
  // 제목과 내용을 조합하여 반환
  return titlePart + "_" + contentPart;
}

/**
 * 텍스트에서 주요 키워드 추출
 * @param {string} text - 원본 텍스트
 * @return {string[]} 추출된 키워드 배열
 */
function extractKeywords(text) {
  if (!text) return [];
  
  // 불용어 목록 - 흔한 한국어 조사, 접속사 등
  const stopwords = [
    "이", "그", "저", "것", "의", "가", "을", "를", "에", "에서", "으로", 
    "와", "과", "이나", "거나", "또는", "및", "에게", "께", "에서", "부터", "까지",
    "이다", "있다", "하다", "되다", "않다", "된다", "한다"
  ];
  
  // 정규화 및 토큰화
  const normalizedText = getNormalizedSignature(text);
  
  // 한글 단어 추출 (2글자 이상)
  const koreanWords = normalizedText.match(/[가-힣]{2,}/g) || [];
  
  // 영어 단어 추출 (3글자 이상)
  const englishWords = normalizedText.match(/[a-z]{3,}/g) || [];
  
  // 한글과 영어 단어 합치기
  const allWords = [...koreanWords, ...englishWords];
  
  // 불용어 제거 및 중복 제거
  return [...new Set(allWords.filter(word => !stopwords.includes(word)))];
}

 /**
 * Gemini API를 활용하여 최종 선택된 뉴스의 중복 여부를 검사하는 함수
 * createAndSendNewsletter 함수 내에서 뉴스레터 생성 직전에 호출
 * @param {Array} topicsWithNews - 주제별로 선택된 뉴스 배열
 * @param {string} googleApiKey - Google AI API 키
 * @return {Array} 중복이 제거된 주제별 뉴스 배열
 */
async function checkFinalDuplicatesWithGemini(topicsWithNews, googleApiKey) {
  try {
    Logger.log("Gemini를 통한 최종 중복 검사 시작");
    
    // 각 주제별로 중복 검사 수행
    const finalTopics = [];
    
    for (const topicObj of topicsWithNews) {
      // 주제 정보 복사
      const newTopicObj = { ...topicObj };
      
      // 국내 뉴스와 해외 뉴스를 분리해서 중복 검사
      if (newTopicObj.domesticNews && newTopicObj.domesticNews.length > 0) {
        newTopicObj.domesticNews = await checkNewsDuplicatesWithGemini(
          newTopicObj.domesticNews, 
          googleApiKey, 
          `${newTopicObj.topic} - 국내 뉴스`
        );
      }
      
      if (newTopicObj.globalNews && newTopicObj.globalNews.length > 0) {
        newTopicObj.globalNews = await checkNewsDuplicatesWithGemini(
          newTopicObj.globalNews, 
          googleApiKey, 
          `${newTopicObj.topic} - 해외 뉴스`
        );
      }
      
      finalTopics.push(newTopicObj);
    }
    
    Logger.log("Gemini를 통한 최종 중복 검사 완료");
    return finalTopics;
    
  } catch (error) {
    Logger.log(`Gemini 중복 검사 중 오류 발생: ${error.message}`);
    // 오류 발생 시 원본 데이터 반환
    return topicsWithNews;
  }
}

/**
 * Gemini API를 사용하여 뉴스 목록의 중복을 검사하는 함수
 * @param {Array} newsList - 뉴스 항목 배열
 * @param {string} apiKey - Google AI API 키
 * @param {string} category - 뉴스 카테고리 (로깅용)
 * @return {Array} 중복이 제거된 뉴스 항목 배열
 */
async function checkNewsDuplicatesWithGemini(newsList, apiKey, category) {
  if (!newsList || newsList.length <= 1) {
    return newsList; // 뉴스가 1개 이하면 중복 검사 불필요
  }
  
  try {
    Logger.log(`${category} 카테고리의 ${newsList.length}개 뉴스 중복 검사 시작`);
    
    // 뉴스 제목과 요약 목록 생성
    const newsInfoList = newsList.map((news, index) => ({
      id: index,
      title: news.title,
      summary: news.aiSummary || news.description || "",
      score: news.aiRelevanceScore || 0
    }));
    
    // Gemini에 중복 검사 요청
    const prompt = createDuplicationCheckPrompt(newsInfoList, category);
    const response = await callGeminiAPIAsync(prompt, apiKey);
    
    // 응답 파싱
    const duplicateGroups = parseGeminiDuplicationResponse(response);
    if (!duplicateGroups || duplicateGroups.length === 0) {
      Logger.log(`${category} - Gemini 응답 파싱 실패, 원본 뉴스 목록 유지`);
      return newsList;
    }
    
    // 중복 그룹에서 대표 뉴스만 선택
    const finalNewsIndices = selectRepresentativeNewsFromGroups(duplicateGroups, newsInfoList);
    
    // 최종 선택된 뉴스 필터링
    const finalNewsList = finalNewsIndices.map(index => newsList[index]);
    
    Logger.log(`${category} - 중복 검사 완료: ${newsList.length}개 중 ${finalNewsList.length}개 선택됨`);
    return finalNewsList;
    
  } catch (error) {
    Logger.log(`${category} 뉴스 중복 검사 중 오류 발생: ${error.message}`);
    return newsList; // 오류 발생 시 원본 뉴스 목록 반환
  }
}

/**
 * 중복 검사를 위한 Gemini 프롬프트 생성
 * @param {Array} newsInfoList - 뉴스 정보 목록
 * @param {string} category - 뉴스 카테고리
 * @return {string} Gemini 프롬프트
 */
function createDuplicationCheckPrompt(newsInfoList, category) {
  const prompt = `다음은 "${category}" 카테고리의 뉴스 목록입니다. 내용이 중복되거나 매우 유사한 뉴스를 그룹화해주세요. 각 그룹에는 하나 이상의 뉴스 ID가 포함됩니다.

내용이 독립적이고 고유한 뉴스는 별도 그룹으로 분류해주세요. 제목이 약간 다르더라도 내용이 매우 유사하면 같은 그룹으로 분류해주세요.

뉴스 목록:
${newsInfoList.map(news => 
  `ID: ${news.id}
제목: ${news.title}
요약: ${news.summary.substring(0, 150)}${news.summary.length > 150 ? '...' : ''}
점수: ${news.score}
`).join('\n---\n')}

다음 JSON 형식으로만 응답해주세요:
[
  {
    "group": 1,
    "newsIds": [0, 2, 5],
    "reason": "포항시 인터배터리 전시회 참가 관련 중복 기사"
  },
  {
    "group": 2,
    "newsIds": [1],
    "reason": "독립적인 기사"
  }
]

JSON 형식만 반환하고 다른 설명이나 주석은 포함하지 마세요.`;

  return prompt;
}

/**
 * Gemini API 비동기 호출 함수
 * @param {string} prompt - 프롬프트
 * @param {string} apiKey - API 키
 * @param {string} modelName - 모델 이름 (옵션)
 * @return {Promise<string>} API 응답 Promise
 */
async function callGeminiAPIAsync(prompt, apiKey, modelName = null) {
  // 모델이 지정되지 않은 경우 기본값 사용
  if (!modelName) {
    modelName = getSelectedGeminiModel();
  }
  
  const apiEndpoint = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent`;
  
  let retryCount = 0;
  const maxRetries = CONFIG.MAX_RETRIES;
  
  while (retryCount < maxRetries) {
    try {
      // 프롬프트 길이 제한 (토큰 제한 방지)
      const truncatedPrompt = prompt.length > 20000 ? prompt.substring(0, 20000) : prompt;
      
      Logger.log(`Gemini API 요청 - 모델: ${modelName}, 프롬프트 길이: ${truncatedPrompt.length}`);
      
      const response = UrlFetchApp.fetch(`${apiEndpoint}?key=${apiKey}`, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{
            parts: [{
              text: truncatedPrompt
            }]
          }],
          generationConfig: {
            temperature: 0.1,
            topP: 0.8,
            topK: 40,
            maxOutputTokens: 2048,
            stopSequences: []
          },
          safetySettings: [
            {
              category: "HARM_CATEGORY_DANGEROUS_CONTENT",
              threshold: "BLOCK_NONE"
            },
            {
              category: "HARM_CATEGORY_HATE_SPEECH",
              threshold: "BLOCK_NONE"
            },
            {
              category: "HARM_CATEGORY_HARASSMENT",
              threshold: "BLOCK_NONE"
            },
            {
              category: "HARM_CATEGORY_SEXUALLY_EXPLICIT",
              threshold: "BLOCK_NONE"
            }
          ]
        }),
        muteHttpExceptions: true
      });
      
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      // 성공 시 응답 반환
      if (responseCode === 200) {
        const responseData = JSON.parse(responseText);
        if (responseData.candidates && responseData.candidates.length > 0 && 
            responseData.candidates[0].content && responseData.candidates[0].content.parts) {
          const textResponse = responseData.candidates[0].content.parts[0].text;
          Logger.log(`Gemini 응답 텍스트 (처음 300자): ${textResponse.substring(0, 300)}...`);
          return textResponse;
        } else {
          Logger.log("Gemini API 응답 형식이 예상과 다릅니다.");
          throw new Error("응답 형식 불일치");
        }
      }
      
      // API 오류 코드에 따른 처리
      if (responseCode === 503 || responseCode === 429) {
        // 서비스 불가(503) 또는 할당량 초과(429)
        retryCount++;
        Logger.log(`Gemini API 호출 실패 (${responseCode}): 재시도 ${retryCount}/${maxRetries}`);
        
        if (retryCount < maxRetries) {
          // 지수 백오프 (재시도마다 대기 시간 증가)
          const waitTime = Math.pow(2, retryCount) * 1000; // 2초, 4초, 8초...
          Utilities.sleep(waitTime);
        }
      } else {
        // 다른 오류는 로그 기록 후 예외 발생
        Logger.log(`Gemini API 오류 (${responseCode}): ${responseText}`);
        throw new Error(`API 오류: ${responseCode}`);
      }
    } catch (error) {
      retryCount++;
      Logger.log(`Gemini API 호출 중 예외 발생: ${error.message}, 재시도 ${retryCount}/${maxRetries}`);
      
      if (retryCount < maxRetries) {
        // 예외 발생 시도 재시도
        const waitTime = Math.pow(2, retryCount) * 1000;
        Utilities.sleep(waitTime);
      } else {
        // 최대 재시도 횟수 초과하면 예외 발생
        throw error;
      }
    }
  }
  
  // 모든 재시도 실패 시 예외 발생
  throw new Error("최대 재시도 횟수 초과");
}

/**
 * Gemini의 중복 검사 응답을 파싱하는 함수
 * @param {string} response - Gemini API 응답
 * @return {Array|null} 중복 그룹 배열 또는 null
 */
function parseGeminiDuplicationResponse(response) {
  try {
    // JSON 형식 추출
    let jsonText = extractJsonFromText(response);
    if (!jsonText) {
      Logger.log("중복 검사 응답에서 JSON을 추출할 수 없습니다.");
      return null;
    }
    
    // JSON 파싱
    const groups = JSON.parse(jsonText);
    
    // 유효성 검사
    if (!Array.isArray(groups)) {
      Logger.log("중복 검사 응답이 배열 형식이 아닙니다.");
      return null;
    }
    
    for (const group of groups) {
      if (!group.group || !Array.isArray(group.newsIds) || !group.reason) {
        Logger.log("중복 그룹 형식이 잘못되었습니다.");
        return null;
      }
    }
    
    return groups;
    
  } catch (error) {
    Logger.log(`중복 검사 응답 파싱 오류: ${error.message}`);
    return null;
  }
}

/**
 * 중복 그룹에서 대표 뉴스를 선택하는 함수
 * @param {Array} duplicateGroups - 중복 그룹 배열
 * @param {Array} newsInfoList - 뉴스 정보 목록
 * @return {Array} 선택된 뉴스 ID 배열
 */
function selectRepresentativeNewsFromGroups(duplicateGroups, newsInfoList) {
  const selectedIndices = [];
  
  for (const group of duplicateGroups) {
    if (group.newsIds.length === 0) continue;
    
    if (group.newsIds.length === 1) {
      // 단일 뉴스 그룹은 그대로 추가
      selectedIndices.push(group.newsIds[0]);
    } else {
      // 다중 뉴스 그룹에서는 가장 높은 점수의 뉴스 선택
      let bestNewsId = group.newsIds[0];
      let highestScore = newsInfoList[bestNewsId].score;
      
      for (let i = 1; i < group.newsIds.length; i++) {
        const currentId = group.newsIds[i];
        const currentScore = newsInfoList[currentId].score;
        
        if (currentScore > highestScore) {
          highestScore = currentScore;
          bestNewsId = currentId;
        }
      }
      
      selectedIndices.push(bestNewsId);
      Logger.log(`중복 그룹 ${group.group}에서 ID ${bestNewsId} 선택 (점수: ${highestScore})`);
    }
  }
  
  return selectedIndices;
}

/**
 * createAndSendNewsletter 함수를 수정하여 Gemini 중복 검사 통합
 * @param {Array} topics - 주제 객체 배열
 * @param {Array} emails - 이메일 주소 배열
 * @param {string} googleApiKey - Google AI API 키
 * @param {string} naverClientId - 네이버 API Client ID
 * @param {string} naverClientSecret - 네이버 API Client Secret
 */
async function createAndSendNewsletterWithGeminiCheck(topics, emails, googleApiKey, naverClientId, naverClientSecret) {
  // 날짜 정보 설정
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  
  const dateInfo = formatDateForNewsletter(today);
  
  // 모든 주제를 하나의 문자열로 결합 (이메일 제목용)
  const topicNames = topics.map(t => t.topic);
  const mainTopic = topicNames.length > 0 ? `${topicNames[0]}/${topicNames.slice(1).join('/')}` : "맞춤 주제";
  
  // 이메일 본문 초기화
  let emailBody = createNewsletterHeader(mainTopic, dateInfo);
  
  // 각 주제별 뉴스 검색 및 추가 (중간 결과 저장)
  const topicsWithNews = [];
  
  let categoryIndex = 0;
  for (const topicObj of topics) {
    // 주제별 뉴스 수집 및 처리
    const { domesticNews, globalNews } = processTopicNewsWithSeparationAndCache(
      topicObj, yesterday, googleApiKey, naverClientId, naverClientSecret
    );
    
    // 주제와 뉴스 정보 저장
    topicsWithNews.push({
      ...topicObj,
      domesticNews,
      globalNews
    });
  }
  
  // Gemini를 통한 최종 중복 검사
  const finalTopics = await checkFinalDuplicatesWithGemini(topicsWithNews, googleApiKey);
  
  // 최종 뉴스로 이메일 본문 생성
  categoryIndex = 0;
  for (const topicWithNews of finalTopics) {
    // 뉴스 카테고리 헤더 생성
    if (categoryIndex > 0) {
      emailBody += `<hr style="border: 0; height: 1px; background-color: #ddd; margin: 25px 0;">`;
    }
    
    emailBody += `<h3 style="color: #1a73e8; margin-top: 20px; margin-bottom: 15px;">■ ${topicWithNews.topic}</h3>`;
    categoryIndex++;
    
    // 국내 뉴스 섹션
    if (topicWithNews.domesticNews && topicWithNews.domesticNews.length > 0) {
      emailBody += `<h4 style="color: #555; margin-top: 15px; margin-bottom: 10px; border-bottom: 1px solid #eee; padding-bottom: 5px;">🇰🇷 국내 뉴스 (${topicWithNews.domesticNews.length}개)</h4>`;
      emailBody += createNewsItemsHtml(topicWithNews.domesticNews, topicWithNews.topic, yesterday, googleApiKey);
    }
    
    // 해외 뉴스 섹션
    if (topicWithNews.globalNews && topicWithNews.globalNews.length > 0) {
      emailBody += `<h4 style="color: #555; margin-top: 20px; margin-bottom: 10px; border-bottom: 1px solid #eee; padding-bottom: 5px;">🌏 해외 뉴스 (${topicWithNews.globalNews.length}개)</h4>`;
      emailBody += createNewsItemsHtml(topicWithNews.globalNews, topicWithNews.topic, yesterday, googleApiKey);
    }
    
    // 뉴스가 없는 경우
    if ((!topicWithNews.domesticNews || topicWithNews.domesticNews.length === 0) && 
        (!topicWithNews.globalNews || topicWithNews.globalNews.length === 0)) {
      emailBody += `<p style="color: #666;">관련 뉴스를 찾을 수 없습니다.</p>`;
    }
  }
  
  // 이메일 닫기 태그
  emailBody += `</div>`;
  
  // 각 이메일 주소로 뉴스레터 발송
  sendEmailToRecipients(emails, mainTopic, dateInfo, emailBody);
}

/**
 * 주제별 뉴스 처리 및 HTML 생성 (국내/해외 뉴스 구분) - 중간 결과 반환
 * @param {Object} topicObj - 주제 객체
 * @param {Date} yesterday - 어제 날짜
 * @param {string} googleApiKey - Google AI API 키
 * @param {string} naverClientId - 네이버 API Client ID
 * @param {string} naverClientSecret - 네이버 API Client Secret
 * @return {Object} 국내 뉴스와 해외 뉴스 객체
 */
function processTopicNewsWithSeparationAndCache(topicObj, yesterday, googleApiKey, naverClientId, naverClientSecret) {
  // 주제 정보 추출
  const { topic, relatedConcepts, includeGlobalNews, englishKeyword, newsApiKey } = topicObj;
  
  // 국내 뉴스와 해외 뉴스 배열 초기화
  let domesticNews = [];
  let globalNews = [];
  
  // 국내 뉴스 검색 및 처리 (네이버 API 사용)
  const koreanNewsItems = searchNaverNews(topic, naverClientId, naverClientSecret);
  if (koreanNewsItems.length > 0) {
    Logger.log(`'${topic}' 주제로 네이버에서 검색된 국내 뉴스: ${koreanNewsItems.length}개`);
    
    // AI 분석 및 필터링
    const filteredLocalNews = filterNewsByGeminiBatched(koreanNewsItems, topic, relatedConcepts, googleApiKey, "국내 뉴스");
    Logger.log(`'${topic}' 주제에 대한 Gemini 분석 후 관련성 높은 국내 뉴스: ${filteredLocalNews.length}개`);
    
    // 국내 뉴스 추가
    if (filteredLocalNews.length > 0) {
      domesticNews = prepareNewsItems(filteredLocalNews, "국내");
    }
  }
  
  // 해외 뉴스 검색 및 처리 (NewsAPI 사용)
  if (includeGlobalNews && newsApiKey) {
    const searchKeyword = englishKeyword || topic;
    const englishNewsItems = searchNewsAPI(searchKeyword, "en", newsApiKey);
    
    if (englishNewsItems.length > 0) {
      Logger.log(`'${topic}' 주제로 NewsAPI에서 해외 뉴스 ${englishNewsItems.length}개 검색됨 (검색 키워드: ${searchKeyword})`);
      
      // AI 분석 및 필터링
      const filteredGlobalNews = filterNewsByGeminiBatched(englishNewsItems, topic, relatedConcepts, googleApiKey, "해외 뉴스");
      Logger.log(`'${topic}' 주제에 대한 Gemini 분석 후 관련성 높은 해외 뉴스: ${filteredGlobalNews.length}개`);
      
      // 해외 뉴스 번역 및 추가
      if (filteredGlobalNews.length > 0) {
        const translatedNews = translateNewsWithGemini(filteredGlobalNews, googleApiKey);
        globalNews = prepareNewsItems(translatedNews, "해외");
      }
    }
  } else if (includeGlobalNews && !newsApiKey) {
    Logger.log(`'${topic}' 주제의 해외 뉴스 검색을 위한 NewsAPI 키가 없습니다.`);
  }
  
  // 각 카테고리별로 관련성 점수로 정렬
  domesticNews.sort((a, b) => b.aiRelevanceScore - a.aiRelevanceScore);
  globalNews.sort((a, b) => b.aiRelevanceScore - a.aiRelevanceScore);
  
  Logger.log(`'${topic}' 주제에 대한 총 관련성 높은 국내 뉴스: ${domesticNews.length}개`);
  Logger.log(`'${topic}' 주제에 대한 총 관련성 높은 해외 뉴스: ${globalNews.length}개`);
  
  return { domesticNews, globalNews };
}


