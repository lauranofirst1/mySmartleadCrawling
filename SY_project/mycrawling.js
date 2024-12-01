const { Actor } = require('apify');
const { PuppeteerCrawler } = require('crawlee');
const xlsx = require('xlsx');
const fs = require('fs');

Actor.main(async () => {
    let sessionCookies;

    const crawler = new PuppeteerCrawler({
        launchContext: {
            launchOptions: {
                headless: false,
                slowMo: 10,
            },
        },
        requestHandlerTimeoutSecs: 120,
        async requestHandler({ page, request }) {
            const studentId = '20235274'; // 학번
            const studentPW = '*******';
            const fileName = `${studentId}_messages.xlsx`;

            // 로그인 처리
            if (!sessionCookies) {
                console.log("로그인중...");
                await page.goto('https://smartlead.hallym.ac.kr/login.php', { waitUntil: 'domcontentloaded' });
                await page.type('input[name="username"]', studentId);
                await page.type('input[name="password"]', studentPW + '\n');
                await page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 30000 });

                if (page.url().includes('login')) { console.log("로그인 실패.");return; }

                console.log("로그인 성공. 쿠키 저장중...");
                sessionCookies = await page.cookies();
            } else {
                console.log("저장된 쿠키 사용중...");
                await page.setCookie(...sessionCookies);
                await page.goto(request.url, { waitUntil: 'networkidle2' });
            }

            // 튜토리얼 건너뛰기 처리
            const tutorialSkipButtonSelector = 'button[data-action="skip"]';
            const tutorialSkipButton = await page.$(tutorialSkipButtonSelector);
            if (tutorialSkipButton) {
                console.log("AI튜터 건너뛰는중...");
                await tutorialSkipButton.click();
                await new Promise(resolve => setTimeout(resolve, 1000));
            } else {
                console.log("AI튜터 건너뛰기 완료.");
            }

            console.log("내 쪽지함으로 이동중...");
            await page.goto('https://smartlead.hallym.ac.kr/local/ubmessage/', { waitUntil: 'networkidle2' });

            console.log("쪽지 리스트 출력중...");
            const messageHistoryExists = await page.waitForSelector('ul.media-list', { timeout: 30000 }).catch(() => null);


            console.log("개별 쪽지함 링크 입력 중...");
            const messageLinks = await page.$$eval('ul.media-list li.media a', links =>
                links
                    .filter(link => !link.textContent.includes('광고') && link.href && !link.href.includes('action.php?type=message_all_delete'))
                    .map(link => link.href)
            );
            console.log("메시지 링크 모음 완료:", messageLinks);

            // 기존 파일 로드 또는 새로 생성
            let workbook;
            if (fs.existsSync(fileName)) {
                workbook = xlsx.readFile(fileName);
            } else {
                workbook = xlsx.utils.book_new();
                workbook.SheetNames.push('Messages');
                workbook.Sheets['Messages'] = xlsx.utils.aoa_to_sheet([['보낸 사람', '받는 사람', '내용', '시간']]);
            }
            const sheet = workbook.Sheets['Messages'];

            // 개별 메시지 내용 수집 및 저장
            let rowIndex = xlsx.utils.sheet_to_json(sheet).length + 2;
            for (const link of messageLinks) {
                console.log(`쪽지함으로 이동: ${link}`);
                await page.goto(link, { waitUntil: 'networkidle2' });

                console.log("쪽지 내용 추출중...");
                const noMessageText = await page.$eval('p.text-danger', el => el.innerText).catch(() => null);
                if (noMessageText && noMessageText.includes('기간동안 등록된 쪽지가 없습니다')) {
                    console.log("열람 기한 만료.");
                    break;
                }

                // 메시지 세부 정보 추출
                const messages = await page.$$eval('div.messages div.message', (messageNodes) => {
                    return messageNodes.map((node) => {
                        // 메시지 방향에 따라 보낸 사람과 받는 사람 구분
                        const isSent = node.classList.contains('to'); // 보낸 메시지인지 확인
                        const senderNode = isSent
                            ? document.querySelector('.fromusers .to_username_link .username') // 보낸 사람 정보
                            : document.querySelector('.tousers .me_username_link .username'); // 받는 사람 정보
                        const receiverNode = isSent
                            ? document.querySelector('.tousers .me_username_link .username') // 받는 사람 정보
                            : document.querySelector('.fromusers .to_username_link .username'); // 보낸 사람 정보

                        const sender = senderNode ? senderNode.textContent.trim() : 'Unknown';
                        const receiver = receiverNode ? receiverNode.textContent.trim() : 'Unknown';

                        // 메시지 내용 추출
                        const contentNode = node.querySelector('div.content');
                        const content = contentNode ? contentNode.innerText.trim() : 'Unknown';

                        // 시간 정보 추출
                        const timeNode = node.querySelector('div.time');
                        const time = timeNode ? timeNode.getAttribute('title') : 'Unknown';

                        return { sender, receiver, content, time };
                    });
                });

                console.log("수집된 쪽지:");
                messages.forEach(({ sender, receiver, content, time }, index) => {
                    console.log(`[쪽지 ${index + 1}]: 보낸 사람: ${sender}, 받는 사람: ${receiver}, 내용: ${content}, 시간: ${time}`);
                    xlsx.utils.sheet_add_aoa(sheet, [[sender, receiver, content, time]], { origin: `A${rowIndex++}` });
                });

                // 빈 줄 추가
                rowIndex++;
            }

            // 파일 저장
            xlsx.writeFile(workbook, fileName);
            console.log(`${fileName} 파일에 데이터 저장 완료.`);
        }
    });

    await crawler.run([{ url: 'https://smartlead.hallym.ac.kr/login.php' }]);
});
