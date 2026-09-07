package kr.co.dss.fx;

import java.util.List;
import java.util.Map;

/**
 * 영업_외자 업무 흐름 0~8단계 (엑셀 '업무관련' 시트 / spec/stages.csv).
 * 단계 정의는 업무 규칙이므로 코드 상수로 두되, 화면 문구는 여기서만 관리한다.
 */
public final class Stages {

    public record Stage(int no, String name, String desc, String store, String tip,
                        List<String> fields, String subfolder, String keyword) {}

    public static final List<Stage> ALL = List.of(
        new Stage(0, "견적서 요청 접수",
            "국내기업(고객사)로부터 메일 또는 전화로 문의 접수",
            "영업관련업무 엑셀시트",
            "최종고객사 · 모델 · 수량 · 필요시기 · 일반/예비기 구분을 빠짐없이 확인하세요.",
            List.of("customer", "customerPic", "endUser", "site", "model", "qty", "kind", "needDate"),
            "02_견적", "견적"),
        new Stage(1, "공급사에 견적내용 전송",
            "최종고객사 / 모델 / 수량 / 필요시기 / 일반·예비기 구분을 정리하여 전송. 견적서 받으면 프린트하여 보관",
            "견적 폴더",
            "공급사에 보낸 내용과 받은 견적서를 자료방에 함께 남겨두세요.",
            List.of("amount", "currency"), "02_견적", "견적"),
        new Stage(2, "발주서(P.O) 접수",
            "국내기업(고객사)로부터 발주서 수령 (외자 SCM 사이트에서 확인 및 프린트). 업무 엑셀시트 기입 및 NAS 외자발주서 폴더 정리",
            "외자발주서 폴더",
            "외자 SCM 사이트에서 P.O 를 확인·프린트하고 발주서 폴더에 정리하세요.",
            List.of("poNo", "poDate", "amount", "currency"), "03_PO", "수주"),
        new Stage(3, "공급사 주문 및 납기 관리",
            "공급사에 P.O / Debit Note 작성 후 함께 발송. 주문서 수령 사인백을 받아 국내기업에 송부. 제품 생산 및 납기 일정 확인, 수시 Follow-up",
            "데빗노트 폴더",
            "FCA 를 입력하면 신용장 기한이 자동 계산됩니다. 사인백 수령 여부를 확인하세요.",
            List.of("debitNo", "offerNo", "fcaDate"), "04_Offer sheet", "발주"),
        new Stage(4, "신용장(L/C) 개설 요청",
            "공급사가 Offer Sheet 송부 → 내용 확인 후 국내기업에 송부. 선적 1주일 전까지 개설요청서 수령 → 일본 Confirm → 응답서(L/C번호)를 공급사에 송부. 수정사항 필요 시 Amend 신청",
            "신용장 폴더",
            "개설요청서는 선적 1주일 전까지 받아야 합니다. 응답서(L/C번호)는 공급사에 송부하세요.",
            List.of("lcRequestDate", "lcBank", "lcNo", "lcOpenDate", "latestShipment", "expiryDate", "lcAmend"),
            "05_신용장", "계약"),
        new Stage(5, "선적서류 수령",
            "공급사에서 Invoice / Packing List / B/L 등 수령. 서류 이상유무 검토, 이상 있을 시 공급사에 수정 요청",
            "각 발주서와 함께 저장",
            "Invoice / Packing List / B/L 이상 유무를 검토하고, 이상 시 공급사에 수정 요청하세요.",
            List.of(), "06_선적", "선적"),
        new Stage(6, "선적서류 국내기업 발송",
            "국내기업(구매처) 담당자에게 선적서류 송부. 필요시 원본/사본 구분하여 전달",
            "각 발주서와 함께 저장",
            "원본/사본 구분이 필요한지 고객사 담당자에게 확인하세요.",
            List.of(), "06_선적", "선적"),
        new Stage(7, "원산지증명서(C/O) 발급",
            "공급사에 원산지증명서 발급 요청 후 선적서류와 함께 제출. 도착 즉시 국내기업에 전달",
            "원산지증명서 폴더",
            "원산지증명서는 도착 즉시 국내기업에 전달합니다.",
            List.of(), "06_선적", "통관"),
        new Stage(8, "커미션 청구 및 수령",
            "모든 납품 완료/검수 후 커미션 청구서를 공급사에 송부 (청구서는 Debit Note 에 있음). 청구한 커미션은 매월 말 25일 전후로 은행을 통해 수령",
            "커미션청구서 폴더",
            "청구서 파일명은 INVOICE DSS<날짜>_Payment Summary for~ 로 시작합니다.",
            List.of("commRate", "commAmount", "commInvoiceNo", "commBilled", "commReceived"),
            "07_정산", "계약")
    );

    public static final int COUNT = ALL.size();
    public static final List<String> STATUSES = List.of("진행중", "보류", "완료", "취소");

    public static Stage get(int no) { return ALL.get(no); }
    public static String name(int no) { return ALL.get(no).name(); }

    /** 단계별 화면 제목(짧은 것) */
    public static final Map<Integer, String> SHORT = Map.of(
        0, "견적요청", 1, "견적전송", 2, "P.O접수", 3, "공급사주문", 4, "신용장",
        5, "서류수령", 6, "서류발송", 7, "원산지", 8, "커미션");

    private Stages() {}
}
