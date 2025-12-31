\
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace PptTplAgent;

internal static class Program
{
    // Office COM 상수 (MsoTriState)
    private const int msoTrue = -1;
    private const int msoFalse = 0;

    private static readonly string LogPath =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "PptTplAgent.log");

    public static int Main(string[] args)
    {
        try
        {
            if (args is null || args.Length == 0) return 0;

            // 레지스트리 "%1"로 넘어온 URL
            var raw = args[0].Trim();

            // 일부 환경에서 따옴표가 포함되어 전달될 수 있음
            raw = raw.Trim('"');

            if (!raw.StartsWith("ppttpl://", StringComparison.OrdinalIgnoreCase))
                return 0;

            var uri = new Uri(raw);

            // 예: ppttpl://insert?src=...&template=...&pos=end
            var cmd = uri.Host.Trim().ToLowerInvariant();

            if (cmd == "ping")
                return 0;

            if (cmd != "insert")
                return 0;

            var qs = ParseQuery(uri.Query);
            if (!qs.TryGetValue("src", out var src) || string.IsNullOrWhiteSpace(src))
                throw new Exception("src 파라미터가 없습니다.");

            qs.TryGetValue("template", out var templateId); // 옵션
            qs.TryGetValue("pos", out var pos);             // end | afterCurrent | at:<n>

            // 1) 구글 링크 → 다운로드 URL로 변환
            var dlUrl = ToDownloadUrl(src);

            // 2) PPTX 다운로드 (임시)
            var tmpPptx = Path.Combine(Path.GetTempPath(), $"ppttpl_{Guid.NewGuid():N}.pptx");
            DownloadFile(dlUrl, tmpPptx).GetAwaiter().GetResult();

            // 3) PPT에 삽입
            InsertPptxIntoPowerPoint(tmpPptx, pos);

            // 4) 정리
            TryDelete(tmpPptx);

            return 0;
        }
        catch (Exception ex)
        {
            SafeLog(ex.ToString());
            return 1;
        }
    }

    // -------------------------
    // Query string 파서 (HttpUtility 없이)
    // -------------------------
    private static Dictionary<string, string> ParseQuery(string query)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (string.IsNullOrWhiteSpace(query))
            return dict;

        var q = query;
        if (q.StartsWith("?")) q = q[1..];

        foreach (var part in q.Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var kv = part.Split('=', 2);
            var key = UrlDecode(kv[0]);
            var val = kv.Length > 1 ? UrlDecode(kv[1]) : "";

            if (!string.IsNullOrWhiteSpace(key))
                dict[key] = val;
        }
        return dict;
    }

    private static string UrlDecode(string s)
    {
        if (s is null) return "";
        s = s.Replace("+", " ");
        return Uri.UnescapeDataString(s);
    }

    // -------------------------
    // Google Slides/Drive URL -> PPTX 다운로드 URL
    // -------------------------
    private static string ToDownloadUrl(string src)
    {
        // 이미 pptx direct 링크면 그대로
        if (src.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
            return src;

        // Google Slides:
        // https://docs.google.com/presentation/d/<ID>/edit...
        var m1 = Regex.Match(src, @"docs\.google\.com\/presentation\/d\/([a-zA-Z0-9_\-]+)");
        if (m1.Success)
        {
            var id = m1.Groups[1].Value;
            return $"https://docs.google.com/presentation/d/{id}/export/pptx";
        }

        // Google Drive file:
        // https://drive.google.com/file/d/<ID>/view...
        var m2 = Regex.Match(src, @"drive\.google\.com\/file\/d\/([a-zA-Z0-9_\-]+)");
        if (m2.Success)
        {
            var id = m2.Groups[1].Value;
            return $"https://drive.google.com/uc?export=download&id={id}";
        }

        // ?id=<ID> 형태
        var m3 = Regex.Match(src, @"[?&]id=([a-zA-Z0-9_\-]+)");
        if (m3.Success)
        {
            var id = m3.Groups[1].Value;
            return $"https://drive.google.com/uc?export=download&id={id}";
        }

        // 그 외는 그대로 시도
        return src;
    }

    private static async System.Threading.Tasks.Task DownloadFile(string url, string path)
    {
        using var handler = new HttpClientHandler
        {
            AllowAutoRedirect = true,
            AutomaticDecompression = System.Net.DecompressionMethods.All
        };

        using var http = new HttpClient(handler)
        {
            Timeout = TimeSpan.FromSeconds(90)
        };

        // Google이 User-Agent에 따라 다른 응답을 줄 때가 있어 기본값 설정
        http.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) PptTplAgent/1.0");

        using var res = await http.GetAsync(url);
        res.EnsureSuccessStatusCode();

        var bytes = await res.Content.ReadAsByteArrayAsync();

        // 간단 방어: 공유 권한이 막혀 있으면 HTML이 떨어지는 경우가 흔함
        var head = Encoding.UTF8.GetString(bytes, 0, Math.Min(bytes.Length, 300));
        if (head.Contains("<html", StringComparison.OrdinalIgnoreCase) ||
            head.Contains("accounts.google.com", StringComparison.OrdinalIgnoreCase))
        {
            throw new Exception(
                "다운로드가 PPTX가 아니라 HTML로 내려왔습니다. (구글 공유 권한이 '링크가 있는 사용자 누구나 보기'인지 확인)");
        }

        await File.WriteAllBytesAsync(path, bytes);
    }

    // -------------------------
    // PowerPoint 삽입
    // -------------------------
    private static void InsertPptxIntoPowerPoint(string pptxPath, string? pos)
    {
        dynamic pptApp = GetOrCreatePowerPoint();

        // ActivePresentation이 없으면 새로 생성
        dynamic? pres = null;
        try { pres = pptApp.ActivePresentation; } catch { /* ignore */ }

        if (pres == null)
        {
            pres = pptApp.Presentations.Add(msoTrue);
        }

        // 소스 PPTX 열어서 슬라이드 수 확인 (창 없이)
        dynamic srcPres = pptApp.Presentations.Open(pptxPath, msoFalse, msoTrue, msoFalse);
        int srcCount = srcPres.Slides.Count;
        srcPres.Close();

        // 삽입 위치 계산
        int insertIndex = pres.Slides.Count; // 기본: 맨 끝 뒤
        pos = (pos ?? "").Trim();

        if (pos.Equals("afterCurrent", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var view = pptApp.ActiveWindow?.View;
                if (view != null)
                {
                    int current = (int)view.Slide.SlideIndex;
                    insertIndex = current; // InsertFromFile는 "Index 뒤에 삽입"
                }
            }
            catch { /* ignore */ }
        }
        else if (pos.StartsWith("at:", StringComparison.OrdinalIgnoreCase))
        {
            // at:3 -> 3번 슬라이드 뒤에 삽입 (Index=3)
            var nStr = pos.Substring(3);
            if (int.TryParse(nStr, out var n) && n >= 0)
            {
                insertIndex = Math.Min(Math.Max(n, 0), (int)pres.Slides.Count);
            }
        }
        // else end(default)

        // InsertFromFile(FileName, Index, SlideStart, SlideEnd)
        pres.Slides.InsertFromFile(pptxPath, insertIndex, 1, srcCount);
    }

    private static dynamic GetOrCreatePowerPoint()
    {
        // 실행 중 PPT에 붙기, 없으면 새로 띄우기
        try
        {
            var active = Marshal.GetActiveObject("PowerPoint.Application");
            if (active != null) return active;
        }
        catch { /* ignore */ }

        var t = Type.GetTypeFromProgID("PowerPoint.Application");
        if (t == null) throw new Exception("PowerPoint가 설치되어 있지 않거나 ProgID를 찾을 수 없습니다.");

        var app = Activator.CreateInstance(t);
        if (app == null) throw new Exception("PowerPoint.Application 인스턴스를 만들지 못했습니다.");

        try { app.Visible = msoTrue; } catch { /* ignore */ }
        return app;
    }

    // -------------------------
    // 유틸
    // -------------------------
    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* ignore */ }
    }

    private static void SafeLog(string msg)
    {
        try
        {
            File.AppendAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}\n");
        }
        catch { /* ignore */ }
    }
}
