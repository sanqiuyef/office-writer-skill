using System.Globalization;
using System.Text.Json;

namespace OfficeNative;

internal static class PptCommands
{
    private static readonly Dictionary<string, int> AnimationEffects = new(StringComparer.OrdinalIgnoreCase)
    {
        ["appear"] = 1,
        ["fly"] = 2,
        ["fly-in"] = 2,
        ["blinds"] = 3,
        ["box"] = 4,
        ["checkerboard"] = 5,
        ["circle"] = 6,
        ["crawl"] = 7,
        ["dissolve"] = 8,
        ["fade"] = 9,
        ["flash-once"] = 10,
        ["peek"] = 11,
        ["plus"] = 12,
        ["random-bars"] = 13,
        ["spiral"] = 14,
        ["split"] = 15,
        ["stretch"] = 16,
        ["strips"] = 17,
        ["swivel"] = 18,
        ["wedge"] = 19,
        ["wheel"] = 20,
        ["wipe"] = 21,
        ["zoom"] = 23,
    };

    private static readonly Dictionary<string, int> TransitionEffects = new(StringComparer.OrdinalIgnoreCase)
    {
        ["none"] = 0,
        ["cut"] = 257,
        ["fade"] = 1793,
        ["dissolve"] = 1537,
        ["blinds-horizontal"] = 769,
        ["blinds-vertical"] = 770,
        ["checkerboard-across"] = 1025,
        ["checkerboard-down"] = 1026,
        ["cover-left"] = 1285,
        ["cover-right"] = 1288,
        ["cover-up"] = 1291,
        ["cover-down"] = 1284,
        ["wipe-left"] = 2817,
        ["wipe-up"] = 2818,
        ["wipe-right"] = 2819,
        ["wipe-down"] = 2820,
    };

    public static Dictionary<string, object?> RunEffects(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var action = NativeCommon.OptString(p, "action") ?? "animations-list";
        NativeCommon.EnsureFile(filePath, ".pptx", ".potx");
        var modifies = action is "animations-add" or "animations-clear"
            or "transition-set" or "transition-clear"
            or "comments-add" or "comments-delete";
        var backupPath = modifies ? NativeCommon.Backup(filePath) : null;
        var session = OpenPresentation(filePath);
        try
        {
            var slide = GetSlide(session.Presentation, p);
            try
            {
                var data = action switch
                {
                    "animations-add" => AddAnimation(slide, p),
                    "animations-list" => ListAnimations(slide),
                    "animations-clear" => ClearAnimations(slide),
                    "transition-set" => SetTransition(slide, p),
                    "transition-clear" => ClearTransition(slide),
                    "transition-list" => ListTransition(slide),
                    "comments-add" => AddComment(slide, p),
                    "comments-list" => ListComments(slide),
                    "comments-delete" => DeleteComments(slide),
                    _ => throw new NativeUsageException(
                        "action 必须是 animations-add/animations-list/animations-clear/"
                        + "transition-set/transition-clear/transition-list/"
                        + "comments-add/comments-list/comments-delete"),
                };
                if (modifies)
                    OpenEditCommon.SavePresentation(session.Presentation);
                data["file_path"] = Path.GetFullPath(filePath);
                if (backupPath is not null)
                    data["backup_path"] = backupPath;
                return data;
            }
            finally
            {
                NativeCommon.ReleaseCom(slide);
            }
        }
        finally
        {
            ClosePresentation(session);
        }
    }

    private static Dictionary<string, object?> AddAnimation(dynamic slide, JsonElement p)
    {
        var shapeIndex = NativeCommon.OptInt(p, "shape_index", 1);
        var effectName = NativeCommon.OptString(p, "effect") ?? "fade";
        var effectId = AnimationEffects.TryGetValue(effectName, out var mapped)
            ? mapped
            : throw new NativeUsageException($"未知动画效果: {effectName}");
        dynamic shape = slide.Shapes.Item(shapeIndex);
        dynamic sequence = slide.TimeLine.MainSequence;
        dynamic effect = sequence.AddEffect(shape, effectId, 0, 1, -1);
        if (p.TryGetProperty("duration", out _))
            effect.Timing.Duration = NativeCommon.OptDouble(p, "duration", 0.5);
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "animations-add",
            ["slide"] = (int)slide.SlideIndex,
            ["shape_index"] = shapeIndex,
            ["shape_name"] = (string)shape.Name,
            ["effect"] = effectName,
            ["effect_count"] = (int)sequence.Count,
        };
    }

    private static Dictionary<string, object?> ListAnimations(dynamic slide)
    {
        dynamic sequence = slide.TimeLine.MainSequence;
        var effects = new List<Dictionary<string, object?>>();
        for (var i = 1; i <= (int)sequence.Count; i++)
        {
            dynamic effect = sequence.Item(i);
            effects.Add(new Dictionary<string, object?>
            {
                ["index"] = i,
                ["shape_name"] = (string)effect.Shape.Name,
                ["effect_type"] = (int)effect.EffectType,
                ["duration"] = (double)effect.Timing.Duration,
            });
        }
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "animations-list",
            ["slide"] = (int)slide.SlideIndex,
            ["count"] = effects.Count,
            ["effects"] = effects,
        };
    }

    private static Dictionary<string, object?> ClearAnimations(dynamic slide)
    {
        dynamic sequence = slide.TimeLine.MainSequence;
        var count = (int)sequence.Count;
        for (var i = count; i >= 1; i--)
            sequence.Item(i).Delete();
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "animations-clear",
            ["slide"] = (int)slide.SlideIndex,
            ["removed"] = count,
        };
    }

    private static Dictionary<string, object?> SetTransition(dynamic slide, JsonElement p)
    {
        var name = NativeCommon.OptString(p, "effect") ?? "fade";
        var entryEffect = TransitionEffects.TryGetValue(name, out var mapped)
            ? mapped
            : NativeCommon.OptInt(p, "entry_effect", -1);
        if (entryEffect < 0)
            throw new NativeUsageException(
                $"未知切换效果: {name}；可传 effect 名称或 entry_effect 数值");
        dynamic transition = slide.SlideShowTransition;
        transition.EntryEffect = entryEffect;
        if (p.TryGetProperty("duration", out _))
            transition.Duration = NativeCommon.OptDouble(p, "duration", 1.0);
        if (p.TryGetProperty("advance_time", out _))
        {
            transition.AdvanceOnTime = -1; // msoTrue
            transition.AdvanceTime = NativeCommon.OptDouble(p, "advance_time", 2.0);
        }
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "transition-set",
            ["slide"] = (int)slide.SlideIndex,
            ["effect"] = name,
            ["entry_effect"] = (int)transition.EntryEffect,
            ["duration"] = (double)transition.Duration,
        };
    }

    private static Dictionary<string, object?> ClearTransition(dynamic slide)
    {
        dynamic transition = slide.SlideShowTransition;
        transition.EntryEffect = 0;
        transition.AdvanceOnTime = 0; // msoFalse
        transition.AdvanceTime = 0;
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "transition-clear",
            ["slide"] = (int)slide.SlideIndex,
        };
    }

    private static Dictionary<string, object?> ListTransition(dynamic slide)
    {
        dynamic transition = slide.SlideShowTransition;
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "transition-list",
            ["slide"] = (int)slide.SlideIndex,
            ["entry_effect"] = (int)transition.EntryEffect,
            ["duration"] = (double)transition.Duration,
            ["advance_on_time"] = (int)transition.AdvanceOnTime == -1,
            ["advance_time"] = (double)transition.AdvanceTime,
        };
    }

    private static Dictionary<string, object?> AddComment(dynamic slide, JsonElement p)
    {
        var text = NativeCommon.RequireString(p, "text");
        var author = NativeCommon.OptString(p, "author") ?? "office-native";
        var initials = NativeCommon.OptString(p, "initials") ?? "ON";
        var x = NativeCommon.OptDouble(p, "x", 100);
        var y = NativeCommon.OptDouble(p, "y", 100);
        dynamic comment = slide.Comments.Add(x, y, author, initials, text);
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "comments-add",
            ["slide"] = (int)slide.SlideIndex,
            ["author"] = (string)comment.Author,
            ["text"] = (string)comment.Text,
            ["comment_count"] = (int)slide.Comments.Count,
        };
    }

    private static Dictionary<string, object?> ListComments(dynamic slide)
    {
        dynamic comments = slide.Comments;
        var items = new List<Dictionary<string, object?>>();
        for (var i = 1; i <= (int)comments.Count; i++)
        {
            dynamic comment = comments.Item(i);
            items.Add(new Dictionary<string, object?>
            {
                ["index"] = i,
                ["author"] = (string)comment.Author,
                ["text"] = (string)comment.Text,
                ["left"] = (double)comment.Left,
                ["top"] = (double)comment.Top,
                ["date"] = comment.DateTime.ToString("o"),
            });
        }
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "comments-list",
            ["slide"] = (int)slide.SlideIndex,
            ["count"] = items.Count,
            ["comments"] = items,
        };
    }

    private static Dictionary<string, object?> DeleteComments(dynamic slide)
    {
        dynamic comments = slide.Comments;
        var count = (int)comments.Count;
        for (var i = count; i >= 1; i--)
            comments.Item(i).Delete();
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "comments-delete",
            ["slide"] = (int)slide.SlideIndex,
            ["removed"] = count,
        };
    }

    // ── SmartArt ─────────────────────────────────────────────────

    public static Dictionary<string, object?> RunSmartArt(JsonElement p)
    {
        var filePath = NativeCommon.RequireString(p, "file_path");
        var action = NativeCommon.OptString(p, "action") ?? "layouts";
        NativeCommon.EnsureFile(filePath, ".pptx", ".potx");
        var backupPath = action == "add" ? NativeCommon.Backup(filePath) : null;
        // SmartArt 布局集合只有在演示文稿带窗口打开时才会加载。
        var session = OpenPresentation(filePath, withWindow: true);
        try
        {
            var data = action switch
            {
                "layouts" => ListSmartArtLayouts(session.App),
                "add" => AddSmartArt(session.App, session.Presentation, p),
                _ => throw new NativeUsageException("action 必须是 layouts/add"),
            };
            if (action == "add")
            {
                OpenEditCommon.SavePresentation(session.Presentation);
                data["backup_path"] = backupPath;
            }
            data["file_path"] = Path.GetFullPath(filePath);
            return data;
        }
        finally
        {
            ClosePresentation(session);
        }
    }

    private static Dictionary<string, object?> ListSmartArtLayouts(dynamic app)
    {
        dynamic layouts = app.SmartArtLayouts;
        var items = new List<Dictionary<string, object?>>();
        for (var i = 1; i <= (int)layouts.Count; i++)
        {
            dynamic layout = layouts.Item(i);
            items.Add(new Dictionary<string, object?>
            {
                ["name"] = (string)layout.Name,
                ["id"] = (string)layout.Id,
            });
        }
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "layouts",
            ["count"] = items.Count,
            ["layouts"] = items,
        };
    }

    private static Dictionary<string, object?> AddSmartArt(dynamic app, dynamic pres, JsonElement p)
    {
        var layoutName = NativeCommon.RequireString(p, "layout");
        var slide = GetSlide(pres, p);
        dynamic layouts = app.SmartArtLayouts;
        if ((int)layouts.Count == 0)
            throw new NativeRuntimeException(
                "SmartArt 布局集合为空；当前 PowerPoint 安装可能缺少 SmartArt 内容，"
                + "或演示文稿未能带窗口加载");
        dynamic? layout = null;
        for (var i = 1; i <= (int)layouts.Count; i++)
        {
            dynamic candidate = layouts.Item(i);
            if (string.Equals((string)candidate.Name, layoutName,
                    StringComparison.OrdinalIgnoreCase)
                || string.Equals((string)candidate.Id, layoutName,
                    StringComparison.OrdinalIgnoreCase))
            {
                layout = candidate;
                break;
            }
        }
        if (layout is null)
            throw new NativeRuntimeException($"找不到 SmartArt 布局: {layoutName}");

        var left = NativeCommon.OptDouble(p, "x", 0.5) * 72;
        var top = NativeCommon.OptDouble(p, "y", 1.5) * 72;
        var width = NativeCommon.OptDouble(p, "width", 8.0) * 72;
        var height = NativeCommon.OptDouble(p, "height", 4.5) * 72;
        dynamic shape = slide.Shapes.AddSmartArt(layout, left, top, width, height);
        dynamic smartArt = shape.SmartArt;
        var texts = NativeCommon.OptStringList(p, "texts");
        var applied = 0;
        if (texts.Count > 0)
        {
            try
            {
                smartArt.AllNodes.Item(1).TextFrame2.TextRange.Text = texts[0];
                applied = 1;
                for (var i = 1; i < texts.Count; i++)
                {
                    dynamic node = smartArt.Nodes.Add();
                    node.TextFrame2.TextRange.Text = texts[i];
                    applied++;
                }
            }
            catch (Exception ex)
            {
                throw new NativeRuntimeException(
                    $"SmartArt 已插入但写入节点失败（布局节点数可能不足）: {ex.Message}");
            }
        }
        return new Dictionary<string, object?>
        {
            ["success"] = true,
            ["action"] = "add",
            ["slide"] = (int)slide.SlideIndex,
            ["shape_name"] = (string)shape.Name,
            ["layout"] = (string)layout.Name,
            ["nodes_written"] = applied,
        };
    }

    // ── COM 生命周期 ─────────────────────────────────────────────

    private sealed class PptComSession
    {
        public required dynamic App { get; init; }
        public required dynamic Presentation { get; init; }
        public bool Started { get; init; }
        public bool OpenedHere { get; init; }
        public int? PreviousDisplayAlerts { get; init; }
    }

    /// <summary>
    /// 优先复用已打开该文件/已运行的 PowerPoint 实例；只有自己新建的实例才 Quit，
    /// 不按 PID 结束进程。共享实例用后恢复 DisplayAlerts。
    /// </summary>
    private static PptComSession OpenPresentation(string filePath, bool withWindow = false)
    {
        var fullPath = Path.GetFullPath(filePath);
        var app = OpenEditCommon.FindPowerPoint(fullPath, out var existing, out _);
        dynamic? presentation = existing;
        var started = false;
        var openedHere = false;
        try
        {
            app ??= GetActivePowerPoint();
            if (app is null)
            {
                var type = Type.GetTypeFromProgID("PowerPoint.Application")
                           ?? throw new NativeRuntimeException("PowerPoint.Application 未注册");
                app = Activator.CreateInstance(type)!;
                started = true;
            }
            if (presentation is null)
            {
                presentation = app.Presentations.Open(
                    fullPath, 0, 0, withWindow ? -1 : 0);
                openedHere = true;
            }
            else if (withWindow)
            {
                try { app.Visible = -1; } catch { /* some builds require visible */ }
            }
            var previousAlerts = GetDisplayAlerts(app);
            SetDisplayAlerts(app, 1); // ppAlertsNone
            return new PptComSession
            {
                App = app,
                Presentation = presentation,
                Started = started,
                OpenedHere = openedHere,
                PreviousDisplayAlerts = previousAlerts,
            };
        }
        catch
        {
            if (started)
            {
                try { app.Quit(); } catch { /* ignore */ }
            }
            NativeCommon.ReleaseCom(presentation);
            NativeCommon.ReleaseCom(app);
            throw;
        }
    }

    private static dynamic? GetActivePowerPoint()
    {
        foreach (var progId in OpenEditCommon.PptProgIds)
        {
            var app = GenericOpenCommands.GetActive(progId);
            if (app is not null)
                return app;
        }
        return null;
    }

    private static int? GetDisplayAlerts(dynamic app)
    {
        try { return (int)app.DisplayAlerts; }
        catch { return null; }
    }

    private static void SetDisplayAlerts(dynamic app, int value)
    {
        try { app.DisplayAlerts = value; }
        catch { /* WPS 或旧版本可能不支持 */ }
    }

    private static dynamic GetSlide(dynamic presentation, JsonElement p)
    {
        var index = NativeCommon.OptInt(p, "slide", 1);
        var count = (int)presentation.Slides.Count;
        if (index < 1 || index > count)
            throw new NativeUsageException($"slide={index} 超出范围（共 {count} 张）");
        return presentation.Slides.Item(index);
    }

    private static void ClosePresentation(PptComSession session)
    {
        if (session.PreviousDisplayAlerts is int previousAlerts)
            SetDisplayAlerts(session.App, previousAlerts);
        if (session.OpenedHere)
        {
            try { session.Presentation.Close(); } catch { /* already closed */ }
        }
        if (session.Started)
        {
            // 释放仍被局部变量/RCW 持有的 COM 对象后再 Quit，避免 PowerPoint 进程残留。
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            try { session.App.Quit(); } catch { /* already exited */ }
        }
        NativeCommon.ReleaseCom(session.Presentation);
        NativeCommon.ReleaseCom(session.App);
    }
}
