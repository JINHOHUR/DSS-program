package kr.co.dss.fx.auth;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpSession;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

/** 현재 요청의 사용자를 어디서든 꺼낸다. 값은 {@link SsoFilter} 가 요청 속성에 넣어둔다. */
public final class UserContext {
    public static final String ATTR = "dss.currentUser";
    public static final String SESSION_DEV = "dss.devUser";

    public static UserInfo current() {
        var attrs = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        if (attrs == null) return UserInfo.ANONYMOUS;
        HttpServletRequest req = attrs.getRequest();
        Object u = req.getAttribute(ATTR);
        if (u instanceof UserInfo ui) return ui;
        HttpSession s = req.getSession(false);
        if (s != null && s.getAttribute(SESSION_DEV) instanceof UserInfo ui) return ui;
        return UserInfo.ANONYMOUS;
    }

    public static String id() {
        String id = current().id();
        return (id == null || id.isBlank()) ? "system" : id;
    }

    private UserContext() {}
}
