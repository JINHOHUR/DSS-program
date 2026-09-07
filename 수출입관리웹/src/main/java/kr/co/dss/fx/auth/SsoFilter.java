package kr.co.dss.fx.auth;

import jakarta.servlet.FilterChain;
import jakarta.servlet.ServletException;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;
import org.springframework.web.filter.OncePerRequestFilter;

import java.io.IOException;
import java.nio.charset.StandardCharsets;

/**
 * SSO 연동 지점.
 *  1) 포털이 헤더(X-User-Id 등)로 사용자를 넘겨주면 그대로 신뢰한다. (포털 앞단에서만 접근 가능하다는 전제)
 *  2) dev-mode 에서는 /dev-login 으로 세션에 넣어둔 사용자를 쓴다.
 *  3) 둘 다 없으면 /dev-login 으로 보낸다(dev-mode) 또는 403.
 * 포털의 실제 연동 방식(토큰/세션)이 정해지면 이 클래스만 바꾸면 된다.
 */
@Component
public class SsoFilter extends OncePerRequestFilter {

    @Value("${dss.sso.header-user:X-User-Id}") private String hUser;
    @Value("${dss.sso.header-name:X-User-Name}") private String hName;
    @Value("${dss.sso.header-role:X-User-Role}") private String hRole;
    @Value("${dss.sso.dev-mode:false}") private boolean devMode;

    @Override
    protected void doFilterInternal(HttpServletRequest req, HttpServletResponse res, FilterChain chain)
            throws ServletException, IOException {
        String path = req.getRequestURI();
        if (path.startsWith("/css/") || path.startsWith("/js/") || path.startsWith("/dev-login")
                || path.equals("/favicon.ico") || path.startsWith("/h2-console")) {
            chain.doFilter(req, res);
            return;
        }

        String uid = req.getHeader(hUser);
        if (uid != null && !uid.isBlank()) {
            req.setAttribute(UserContext.ATTR, new UserInfo(uid.trim(),
                    decode(req.getHeader(hName)), decode(req.getHeader(hRole))));
            chain.doFilter(req, res);
            return;
        }

        HttpSession s = req.getSession(false);
        if (s != null && s.getAttribute(UserContext.SESSION_DEV) != null) {
            chain.doFilter(req, res);
            return;
        }

        if (devMode) {
            res.sendRedirect("/dev-login?next=" + java.net.URLEncoder.encode(
                    path + (req.getQueryString() != null ? "?" + req.getQueryString() : ""), StandardCharsets.UTF_8));
        } else {
            res.sendError(403, "포털을 통해 접속하세요.");
        }
    }

    /** 헤더에 한글이 오면 보통 UTF-8 percent-encoding 이나 RFC 2047 이다. 둘 다 아니면 원문. */
    private static String decode(String v) {
        if (v == null) return "";
        try {
            return java.net.URLDecoder.decode(v, StandardCharsets.UTF_8);
        } catch (Exception e) {
            return v;
        }
    }
}
