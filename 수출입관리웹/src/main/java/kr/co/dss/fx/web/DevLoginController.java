package kr.co.dss.fx.web;

import jakarta.servlet.http.HttpSession;
import kr.co.dss.fx.auth.UserContext;
import kr.co.dss.fx.auth.UserInfo;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;

/** 포털 연동 전 개발용 임시 로그인. dev-mode=false 면 404. */
@Controller
public class DevLoginController {
    @Value("${dss.sso.dev-mode:false}") private boolean devMode;

    @GetMapping("/dev-login")
    public String form(@RequestParam(defaultValue = "/") String next, Model m) {
        if (!devMode) return "redirect:/";
        m.addAttribute("next", next);
        return "dev-login";
    }

    @PostMapping("/dev-login")
    public String login(@RequestParam String user, @RequestParam(defaultValue = "") String name,
                        @RequestParam(defaultValue = "영업관리") String role,
                        @RequestParam(defaultValue = "/") String next, HttpSession session) {
        if (!devMode) return "redirect:/";
        session.setAttribute(UserContext.SESSION_DEV, new UserInfo(user.trim(), name.trim(), role));
        return "redirect:" + (next.startsWith("/") ? next : "/");
    }

    @GetMapping("/logout")
    public String logout(HttpSession session) { session.invalidate(); return "redirect:/dev-login"; }
}
