const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle } = require('docx');

// 공통 셀 생성 헬퍼 함수
function createCell(text, bold = false, align = AlignmentType.LEFT, widthPercent = null, fill = null, colSpan = null, rowSpan = null) {
    const runs = [];
    if (typeof text === 'string') {
        const lines = text.split('\n');
        lines.forEach((line, i) => {
            if (i > 0) runs.push(new TextRun({ break: 1 }));
            runs.push(new TextRun({ text: line, bold: bold, font: "맑은 고딕", size: 19 }));
        });
    }

    const cellOptions = {
        children: [new Paragraph({ children: runs, alignment: align })],
        margins: { top: 120, bottom: 120, left: 150, right: 150 },
        verticalAlign: "center"
    };

    if (widthPercent) {
        cellOptions.width = { size: widthPercent * 50, type: WidthType.PERCENTAGE };
    }
    if (fill) {
        cellOptions.shading = { fill: fill };
    }
    if (colSpan) {
        cellOptions.columnSpan = colSpan;
    }
    if (rowSpan) {
        cellOptions.rowSpan = rowSpan;
    }
    
    // 테두리 설정 (연한 회색 테두리)
    cellOptions.borders = {
        top: { style: BorderStyle.SINGLE, size: 4, color: "cbd5e1" },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: "cbd5e1" },
        left: { style: BorderStyle.SINGLE, size: 4, color: "cbd5e1" },
        right: { style: BorderStyle.SINGLE, size: 4, color: "cbd5e1" }
    };

    return new TableCell(cellOptions);
}

// DOCX 파일 생성 공통 함수
function generateDocument(title, roleName, outPath) {
    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                // 문서 대제목
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 400 },
                    children: [
                        new TextRun({
                            text: title,
                            bold: true,
                            size: 32,
                            font: "맑은 고딕",
                        })
                    ]
                }),
                
                // 1. 제품/서비스 개요
                new Paragraph({
                    spacing: { before: 200, after: 100 },
                    children: [
                        new TextRun({
                            text: "□ 제품/서비스 개요",
                            bold: true,
                            size: 24,
                            font: "맑은 고딕",
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createCell("제품/서비스명", true, AlignmentType.CENTER, 15, "f1f5f9"),
                                createCell("학습용 웹앱 포털", false, AlignmentType.LEFT, 35),
                                createCell(roleName === "eduzip" ? "공급자" : "검토자", true, AlignmentType.CENTER, 15, "f1f5f9"),
                                createCell("우신중학교 교사 홍민아", false, AlignmentType.LEFT, 35),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("접속경로", true, AlignmentType.CENTER, 15, "f1f5f9"),
                                createCell("https://minah-s-digital-platform.web.app", false, AlignmentType.LEFT, 85, null, 3),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("주요 내용 및\n기능 · 특장점", true, AlignmentType.CENTER, 15, "f1f5f9"),
                                createCell(
                                    "본 서비스는 우신중학교 교사 홍민아가 개발한 학급 수업 및 학업 지원용 '학습용 웹앱 포털'입니다. 교육 활동에 필요한 다양한 웹 애플리케이션들을 한곳에서 접근하고 관리할 수 있도록 제공하며, 별도의 회원가입 없이 브라우저 로컬 저장소(LocalStorage)에 학생 정보 등을 안전하게 임시 보관하므로 개인정보 유출 우려가 전혀 없는 고효율 안전 학습 도구입니다.",
                                    false, AlignmentType.LEFT, 85, null, 3
                                ),
                            ]
                        })
                    ]
                }),

                // 2. 개인정보보호 기준 충족여부
                new Paragraph({
                    spacing: { before: 400, after: 100 },
                    children: [
                        new TextRun({
                            text: "□ 개인정보보호 기준 충족여부",
                            bold: true,
                            size: 24,
                            font: "맑은 고딕",
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createCell("선정기준", true, AlignmentType.CENTER, 15, "f1f5f9"),
                                createCell("세부 내용", true, AlignmentType.CENTER, 35, "f1f5f9"),
                                createCell("확인", true, AlignmentType.CENTER, 25, "f1f5f9", 3),
                                createCell("증빙 및 근거 (약관 조항)", true, AlignmentType.CENTER, 25, "f1f5f9"),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("", false, AlignmentType.CENTER, 15, "f1f5f9"),
                                createCell("", false, AlignmentType.CENTER, 35, "f1f5f9"),
                                createCell("충족", true, AlignmentType.CENTER, 8, "f1f5f9"),
                                createCell("미충족", true, AlignmentType.CENTER, 8, "f1f5f9"),
                                createCell("해당없음", true, AlignmentType.CENTER, 9, "f1f5f9"),
                                createCell("", false, AlignmentType.CENTER, 25, "f1f5f9"),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("1. 최소처리\n원칙 준수", true, AlignmentType.CENTER, 15, null, null, 3),
                                createCell("1-1. 개인정보가 최소한으로 수집되는가?", false, AlignmentType.LEFT, 35),
                                createCell("■", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 9),
                                createCell("(개인정보 처리방침 제3조)\n수업 진행 및 추첨 기능을 수행하기 위한 최소한의 데이터(학생 이름/닉네임)만 입력받으며 불필요한 개인정보(비밀번호, 연락처 등)는 일체 요구하지 않음", false, AlignmentType.LEFT, 25),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("1-2. 개인정보 수집·이용 목적이 기재되어 있는가?", false, AlignmentType.LEFT, 35),
                                createCell("■", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 9),
                                createCell("(개인정보 처리방침 제1조)\n학생 명단 식별 및 추첨 이력 관리를 목적으로 최소한의 데이터만 처리함을 약관상 명문화하여 고지함", false, AlignmentType.LEFT, 25),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("1-3. 개인정보 수집항목, 보유기간 등이 기재되어 있는가?", false, AlignmentType.LEFT, 35),
                                createCell("■", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 9),
                                createCell("(개인정보 처리방침 제2조, 제3조)\n로컬 스토리지에만 저장되는 데이터 수집 항목 및 브라우저 캐시 삭제 시 파기되는 보유기간 명시", false, AlignmentType.LEFT, 25),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("2. 개인정보\n안전조치 의무", true, AlignmentType.CENTER, 15),
                                createCell("2-1. 개인정보 안전성 확보에 필요한 조치 사항이 기재되어 있는가?", false, AlignmentType.LEFT, 35),
                                createCell("■", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 9),
                                createCell("(개인정보 처리방침 제6조)\n서버 수집 배제(해킹 유출 차단) 및 보안전송 프로토콜(HTTPS) 적용을 통한 전송구간 안전 조치 기재", false, AlignmentType.LEFT, 25),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("3. 열람/정정/\n삭제/처리정지\n절차", true, AlignmentType.CENTER, 15),
                                createCell("3-1. 이용자에게 언제든지 자신의 정보를 열람·정정·삭제·처리정지를 요구할 수 있는 절차가 안내되어 있는가?", false, AlignmentType.LEFT, 35),
                                createCell("■", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 9),
                                createCell("(개인정보 처리방침 제7조)\n기기 내 로컬 스토리지 방식을 통해 이용자가 화면에서 직접 수정 및 [초기화] 버튼으로 즉시 영구 삭제 가능한 통제권 부여함", false, AlignmentType.LEFT, 25),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("4. 만14세 미만\n아동의\n개인정보 보호", true, AlignmentType.CENTER, 15),
                                createCell("4-1. 만 14세 미만 아동의 경우 법정대리인 동의 등 아동의 개인정보 보호를 위한 절차가 마련되어 있는가?", false, AlignmentType.LEFT, 35),
                                createCell("■", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 9),
                                createCell("(개인정보 처리방침 제4조)\n서버 DB 저장이 없어 외부 유출 우려가 없으며, 필요한 경우 법정대리인의 개인정보 수집 이용 동의 가이드라인 제시", false, AlignmentType.LEFT, 25),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("5. 보호책임자/\n제3자제공/\n위탁 등", true, AlignmentType.CENTER, 15, null, null, null, 3),
                                createCell("5-1. 개인정보 보호책임자 관련 정보가 안내되어 있는가?", false, AlignmentType.LEFT, 35),
                                createCell("■", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 9),
                                createCell("(개인정보 처리방침 제8조)\n개인정보 보호책임자로 개발자 본인(교사 홍민아, 우신중학교, 02-2610-1621) 연락정보 기재", false, AlignmentType.LEFT, 25),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("5-2. 개인정보 제3자 제공에 관한 정보가 기재되어 있는가? (필요시)", false, AlignmentType.LEFT, 35),
                                createCell("□", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 8),
                                createCell("■", false, AlignmentType.CENTER, 9),
                                createCell("(해당없음)\n클라이언트 사이드 전용 앱으로 개인정보를 제3자에게 수집하거나 제공하지 않음", false, AlignmentType.LEFT, 25),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("5-3. 개인정보 위·수탁관계에 관한 정보가 기재되어 있는가? (필요시)", false, AlignmentType.LEFT, 35),
                                createCell("□", false, AlignmentType.CENTER, 8),
                                createCell("□", false, AlignmentType.CENTER, 8),
                                createCell("■", false, AlignmentType.CENTER, 9),
                                createCell("(해당없음)\n개인정보 수집 및 처리 프로세스를 외부 업체나 제3자에 위탁하거나 수탁 관계를 맺고 있지 않음", false, AlignmentType.LEFT, 25),
                            ]
                        })
                    ]
                }),

                // 3. 작성일 및 문의처
                new Paragraph({
                    spacing: { before: 400, after: 100 },
                    children: [
                        new TextRun({
                            text: "□ 작성일 및 문의처",
                            bold: true,
                            size: 24,
                            font: "맑은 고딕",
                        })
                    ]
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                createCell("작성일", true, AlignmentType.CENTER, 15, "f1f5f9"),
                                createCell("2026. 06. 27.", false, AlignmentType.LEFT, 85),
                            ]
                        }),
                        new TableRow({
                            children: [
                                createCell("문의처", true, AlignmentType.CENTER, 15, "f1f5f9"),
                                createCell("우신중학교 교사 홍민아 (02-2610-1621 / minah2554@users.noreply.github.com)", false, AlignmentType.LEFT, 85),
                            ]
                        })
                    ]
                })
            ]
        }]
    });

    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync(outPath, buffer);
        console.log(`성공적으로 생성되었습니다: ${outPath}`);
    });
}

// 1. 에듀집 탑재용 체크리스트 생성
generateDocument(
    "학습지원 소프트웨어 필수기준 체크리스트 (에듀집 탑재용)",
    "eduzip",
    "c:\\Users\\user\\Desktop\\웹 포털 제작\\에듀집_탑재용_필수기준_체크리스트_학습용_웹앱_포털.docx"
);

// 2. 학교용 필수기준 체크리스트 생성
generateDocument(
    "학습지원 소프트웨어 필수기준 체크리스트 (학교 자체 검토용)",
    "school",
    "c:\\Users\\user\\Desktop\\웹 포털 제작\\학교용_필수기준_체크리스트_학습용_웹앱_포털.docx"
);
