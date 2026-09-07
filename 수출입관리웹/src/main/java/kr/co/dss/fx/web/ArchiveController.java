package kr.co.dss.fx.web;

import kr.co.dss.fx.service.ArchiveService;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

@Controller
@RequestMapping("/archive")
public class ArchiveController {
    private final ArchiveService archive;
    public ArchiveController(ArchiveService archive) { this.archive = archive; }

    @GetMapping
    public String browse(@RequestParam(defaultValue = "") String path, @RequestParam(defaultValue = "") String q, Model m) {
        m.addAttribute("root", archive.rootText());
        m.addAttribute("rootOk", archive.rootExists());
        m.addAttribute("tops", archive.topFolders());
        m.addAttribute("path", path);
        m.addAttribute("q", q);
        if (!q.isBlank()) m.addAttribute("results", archive.search(path, q));
        m.addAttribute("listing", archive.list(path));
        m.addAttribute("active_menu", "archive");
        return "archive";
    }

    @GetMapping("/download")
    public ResponseEntity<FileSystemResource> download(@RequestParam String path) throws Exception {
        Path p = archive.resolve(path);
        if (!Files.isRegularFile(p)) return ResponseEntity.notFound().build();
        String name = p.getFileName().toString();
        String enc = URLEncoder.encode(name, StandardCharsets.UTF_8).replace("+", "%20");
        String ct = Files.probeContentType(p);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "inline; filename*=UTF-8''" + enc)
                .contentType(ct == null ? MediaType.APPLICATION_OCTET_STREAM : MediaType.parseMediaType(ct))
                .body(new FileSystemResource(p));
    }

    @PostMapping("/upload")
    public String upload(@RequestParam String path, @RequestParam("files") MultipartFile[] files,
                         @RequestParam(defaultValue = "rename") String conflict, RedirectAttributes ra) {
        try {
            var msgs = archive.upload(path, files, conflict);
            ra.addFlashAttribute("msg", msgs.size() + "개 처리: " + String.join(", ", msgs.subList(0, Math.min(5, msgs.size()))) + (msgs.size() > 5 ? " …" : ""));
        } catch (Exception e) { ra.addFlashAttribute("err", e.getMessage()); }
        return back(path);
    }

    @PostMapping("/mkdir")
    public String mkdir(@RequestParam String path, @RequestParam String name, RedirectAttributes ra) {
        try { archive.mkdir(path, name); ra.addFlashAttribute("msg", "폴더 생성: " + name); }
        catch (Exception e) { ra.addFlashAttribute("err", e.getMessage()); }
        return back(path);
    }

    @PostMapping("/rename")
    public String rename(@RequestParam String path, @RequestParam String target, @RequestParam String name, RedirectAttributes ra) {
        try { archive.rename(target, name); ra.addFlashAttribute("msg", "이름 변경: " + name); }
        catch (Exception e) { ra.addFlashAttribute("err", e.getMessage()); }
        return back(path);
    }

    @PostMapping("/delete")
    public String delete(@RequestParam String path, @RequestParam("targets") java.util.List<String> targets, RedirectAttributes ra) {
        int ok = 0;
        for (String t : targets) { try { archive.trash(t); ok++; } catch (Exception ignore) {} }
        ra.addFlashAttribute("msg", ok + "개를 휴지통(" + ArchiveService.TRASH + ")으로 옮겼습니다");
        return back(path);
    }

    private static String back(String path) {
        return "redirect:/archive?path=" + URLEncoder.encode(path == null ? "" : path, StandardCharsets.UTF_8);
    }
}
