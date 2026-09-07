package kr.co.dss.fx.auth;

/** 포털(SSO)이 넘겨주는 사용자. 이 모듈은 비밀번호·로그인 화면을 갖지 않는다. */
public record UserInfo(String id, String name, String role) {
    public static final UserInfo ANONYMOUS = new UserInfo("", "", "");

    public boolean isPresent() { return id != null && !id.isBlank(); }
    public boolean canEdit() {
        // 포털 역할 체계(영업관리·수리팀·회계팀·관리자·대표): 조회 전용 역할만 막는다
        return !"조회".equals(role) && !"viewer".equalsIgnoreCase(role);
    }
    public String display() { return (name == null || name.isBlank()) ? id : name; }
}
