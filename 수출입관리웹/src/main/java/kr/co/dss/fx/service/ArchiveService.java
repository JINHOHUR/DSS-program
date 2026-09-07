package kr.co.dss.fx.service;

import kr.co.dss.fx.Stages;
import kr.co.dss.fx.domain.Deal;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Stream;

/**
 * 자료방 — 설정된 루트(로컬 또는 NAS UNC) 아래를 탐색·업로드·정리한다.
 * 경로는 항상 루트 기준 상대경로("04. 견적/001_외자견적서")로 주고받고, 루트 밖 접근은 막는다.
 * 삭제는 루트의 "_휴지통" 으로 옮긴다(복구 가능).
 */
@Service
public class ArchiveService {
    private final SettingService settings;
    public static final String TRASH = "_휴지통";
    private static final DateTimeFormatter TS = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm");

    public ArchiveService(SettingService settings) { this.settings = settings; }

    public record Entry(String name, String rel, boolean dir, long size, String mtime, int count, String kind) {
        public String sizeText() { return dir ? count + "개" : ArchiveService.size(size); }
    }
    public record Listing(String rel, List<String> crumbs, List<Entry> entries, int dirs, int files, long total, boolean ok, String error) {
        public String totalText() { return ArchiveService.size(total); }
    }

    public Path root() {
        String r = settings.archiveRoot();
        return (r == null || r.isBlank()) ? null : Paths.get(r).toAbsolutePath().normalize();
    }
    public boolean rootExists() { Path r = root(); return r != null && Files.isDirectory(r); }
    public String rootText() { Path r = root(); return r == null ? "" : r.toString(); }

    /** 상대경로 → 절대경로 (루트 밖이면 예외) */
    public Path resolve(String rel) {
        Path root = root();
        if (root == null) throw new IllegalStateException("자료방 경로가 설정되지 않았습니다.");
        Path p = (rel == null || rel.isBlank() || ".".equals(rel)) ? root : root.resolve(rel.replace('\\', '/')).normalize();
        if (!p.startsWith(root)) throw new IllegalArgumentException("자료방 밖의 경로입니다.");
        return p;
    }
    public String rel(Path p) {
        Path root = root();
        String s = root.relativize(p.toAbsolutePath().normalize()).toString().replace('\\', '/');
        return s;
    }

    public List<String> topFolders() {
        Path root = root();
        if (root == null || !Files.isDirectory(root)) return List.of();
        try (Stream<Path> st = Files.list(root)) {
            return st.filter(Files::isDirectory).map(p -> p.getFileName().toString())
                    .filter(n -> !n.startsWith("_") && !n.startsWith("#")).sorted().toList();
        } catch (IOException e) { return List.of(); }
    }

    public Listing list(String rel) {
        Path dir;
        try { dir = resolve(rel); } catch (RuntimeException e) { return new Listing(rel, List.of(), List.of(), 0, 0, 0, false, e.getMessage()); }
        if (!Files.isDirectory(dir)) return new Listing(rel, List.of(), List.of(), 0, 0, 0, false, "폴더가 없습니다: " + rel);
        List<Entry> dirs = new ArrayList<>(), files = new ArrayList<>();
        long total = 0;
        try (Stream<Path> st = Files.list(dir)) {
            for (Path p : st.sorted(Comparator.comparing(x -> x.getFileName().toString(), String.CASE_INSENSITIVE_ORDER)).toList()) {
                String name = p.getFileName().toString();
                if (name.startsWith("#") || name.equalsIgnoreCase("Thumbs.db")) continue;
                if (Files.isDirectory(p)) {
                    int cnt = 0;
                    try (Stream<Path> c = Files.list(p)) { cnt = (int) c.count(); } catch (IOException ignore) {}
                    dirs.add(new Entry(name, rel(p), true, 0, mtime(p), cnt, "폴더"));
                } else {
                    long sz = 0; try { sz = Files.size(p); } catch (IOException ignore) {}
                    total += sz;
                    files.add(new Entry(name, rel(p), false, sz, mtime(p), 0, kind(name)));
                }
            }
        } catch (IOException e) { return new Listing(rel, List.of(), List.of(), 0, 0, 0, false, e.getMessage()); }
        List<Entry> all = new ArrayList<>(dirs); all.addAll(files);
        return new Listing(rel(dir), crumbs(rel(dir)), all, dirs.size(), files.size(), total, true, null);
    }

    public List<Entry> search(String rel, String q) {
        Path base;
        try { base = resolve(rel); } catch (RuntimeException e) { return List.of(); }
        String qq = q.toLowerCase();
        List<Entry> out = new ArrayList<>();
        try (Stream<Path> st = Files.walk(base)) {
            for (Path p : (Iterable<Path>) st::iterator) {
                if (out.size() >= 500) break;
                String name = p.getFileName() == null ? "" : p.getFileName().toString();
                if (Files.isRegularFile(p) && name.toLowerCase().contains(qq) && !name.equalsIgnoreCase("Thumbs.db")) {
                    long sz = 0; try { sz = Files.size(p); } catch (IOException ignore) {}
                    out.add(new Entry(name, rel(p), false, sz, mtime(p), 0, rel(p.getParent())));
                }
            }
        } catch (IOException ignore) {}
        return out;
    }

    // ---- 쓰기 ------------------------------------------------------------------
    /** 업로드. conflict: rename | overwrite | skip. 반환: 결과 메시지 목록 */
    public List<String> upload(String rel, MultipartFile[] files, String conflict) throws IOException {
        Path dir = resolve(rel);
        if (!Files.isDirectory(dir)) throw new IllegalArgumentException("폴더가 없습니다.");
        List<String> msgs = new ArrayList<>();
        for (MultipartFile f : files) {
            if (f.isEmpty()) continue;
            String name = Paths.get(f.getOriginalFilename() == null ? "file" : f.getOriginalFilename()).getFileName().toString();
            Path dst = dir.resolve(name);
            if (Files.exists(dst)) {
                if ("skip".equals(conflict)) { msgs.add("건너뜀: " + name); continue; }
                if (!"overwrite".equals(conflict)) dst = unique(dir, name);
            }
            try (var in = f.getInputStream()) { Files.copy(in, dst, StandardCopyOption.REPLACE_EXISTING); }
            msgs.add("올림: " + dst.getFileName());
        }
        return msgs;
    }

    public Path mkdir(String rel, String name) throws IOException {
        check(name);
        Path p = resolve(rel).resolve(name);
        if (Files.exists(p)) throw new IllegalArgumentException("이미 같은 이름이 있습니다.");
        return Files.createDirectories(p);
    }

    public Path rename(String rel, String newName) throws IOException {
        check(newName);
        Path src = resolve(rel);
        Path dst = src.resolveSibling(newName);
        if (Files.exists(dst)) throw new IllegalArgumentException("이미 같은 이름이 있습니다.");
        return Files.move(src, dst);
    }

    /** 휴지통(_휴지통/yyyyMMdd_HHmmss_이름)으로 이동 */
    public Path trash(String rel) throws IOException {
        Path src = resolve(rel);
        Path trash = root().resolve(TRASH);
        Files.createDirectories(trash);
        String stamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        return Files.move(src, trash.resolve(stamp + "_" + src.getFileName()), StandardCopyOption.REPLACE_EXISTING);
    }

    // ---- 단계 ↔ 폴더 ----------------------------------------------------------
    /** 안건·단계에 해당하는 자료방 상대경로. 없으면 "" (루트) */
    public String stageFolder(Deal d, Integer stageNo) {
        Path root = root();
        if (root == null || !Files.isDirectory(root)) return "";
        String ad = d == null ? null : d.getArchiveDir();
        Path base = null;
        if (ad != null && !ad.isBlank()) {
            try { base = resolve(ad); } catch (RuntimeException ignore) {}
            if (base != null && !Files.isDirectory(base)) base = null;
        }
        if (stageNo == null) return base == null ? "" : rel(base);
        Stages.Stage st = Stages.get(stageNo);
        if (base != null) {
            try (Stream<Path> s = Files.list(base)) {
                Optional<Path> hit = s.filter(Files::isDirectory)
                        .filter(p -> p.getFileName().toString().equalsIgnoreCase(st.subfolder())).findFirst();
                if (hit.isPresent()) return rel(hit.get());
            } catch (IOException ignore) {}
            return rel(base);
        }
        return keywordFolder(st.keyword());
    }

    /** 자료방 최상위에서 키워드를 포함하는 첫 폴더 (정렬순) */
    public String keywordFolder(String keyword) {
        if (keyword == null) return "";
        for (String n : topFolders()) if (n.contains(keyword)) return n;
        return "";
    }

    // ---- 유틸 ------------------------------------------------------------------
    private static Path unique(Path dir, String name) {
        int dot = name.lastIndexOf('.');
        String base = dot > 0 ? name.substring(0, dot) : name, ext = dot > 0 ? name.substring(dot) : "";
        for (int i = 2; ; i++) {
            Path p = dir.resolve(base + " (" + i + ")" + ext);
            if (!Files.exists(p)) return p;
        }
    }
    private static void check(String name) {
        if (name == null || name.isBlank() || name.matches(".*[\\\\/:*?\"<>|].*") || name.startsWith("."))
            throw new IllegalArgumentException("이름에 \\ / : * ? \" < > | 는 쓸 수 없습니다.");
    }
    private static List<String> crumbs(String rel) {
        List<String> out = new ArrayList<>();
        if (rel == null || rel.isBlank()) return out;
        StringBuilder acc = new StringBuilder();
        for (String part : rel.split("/")) {
            if (part.isBlank()) continue;
            if (acc.length() > 0) acc.append('/');
            acc.append(part);
            out.add(acc.toString());
        }
        return out;
    }
    private static String mtime(Path p) {
        try {
            BasicFileAttributes a = Files.readAttributes(p, BasicFileAttributes.class);
            return LocalDateTime.ofInstant(a.lastModifiedTime().toInstant(), ZoneId.systemDefault()).format(TS);
        } catch (IOException e) { return ""; }
    }
    public static String size(long n) {
        if (n < 1024) return n + " B";
        if (n < 1024 * 1024) return (n / 1024) + " KB";
        if (n < 1024L * 1024 * 1024) return String.format("%.1f MB", n / 1048576.0);
        return String.format("%.1f GB", n / 1073741824.0);
    }
    public static String kind(String name) {
        String ext = name.contains(".") ? name.substring(name.lastIndexOf('.') + 1).toLowerCase() : "";
        return switch (ext) {
            case "xlsx", "xlsm", "xls" -> "Excel"; case "csv" -> "CSV";
            case "docx", "doc" -> "Word"; case "pptx", "ppt" -> "PowerPoint"; case "pdf" -> "PDF";
            case "hwp", "hwpx" -> "한글"; case "msg", "eml" -> "메일"; case "txt" -> "텍스트";
            case "png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff" -> "이미지";
            case "zip", "7z", "rar", "alz" -> "압축"; case "dwg" -> "도면"; case "lnk" -> "바로가기";
            default -> "파일";
        };
    }
}
