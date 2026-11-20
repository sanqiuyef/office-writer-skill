namespace OfficeNative;

/// <summary>单项检查结果；字段与 Python checker.CheckResult / to_dict() 对齐。</summary>
internal sealed class WordValidateCheck
{
    public string Id { get; init; } = "";
    public string Name { get; init; } = "";
    public bool Passed { get; init; }
    public string Severity { get; init; } = "info";
    public string Message { get; init; } = "";
    public string Details { get; init; } = "";
}

/// <summary>
/// 验证报告；端口 checker.ValidationReport。checks 保持 Python 的顺序与计数语义：
/// passed 才是 passed，未通过且 severity=error 计入 errors，其余未通过计入 warnings。
/// </summary>
internal sealed class WordValidateReport
{
    private readonly List<WordValidateCheck> _checks = [];

    public WordValidateReport(string filePath) => FilePath = filePath;

    public string FilePath { get; }
    public IReadOnlyList<WordValidateCheck> Checks => _checks;
    public int Total { get; private set; }
    public int Passed { get; private set; }
    public int Warnings { get; private set; }
    public int Errors { get; private set; }

    /// <summary>severity=info 的检查数（与 passed 等值，保留 Python 之外约定的汇总字段）。</summary>
    public int Infos => _checks.Count(check => check.Severity == "info");

    public void Add(
        string name, bool passed, string severity = "info", string message = "",
        string details = "", string checkId = "")
    {
        _checks.Add(new WordValidateCheck
        {
            Id = string.IsNullOrEmpty(checkId) ? name : checkId,
            Name = name,
            Passed = passed,
            Severity = severity,
            Message = message,
            Details = details,
        });
        Total++;
        if (passed)
            Passed++;
        else if (severity == "error")
            Errors++;
        else
            Warnings++;
    }

    public Dictionary<string, object?> ToDictionary()
    {
        var checks = _checks.Select(check => new Dictionary<string, object?>
        {
            ["id"] = check.Id,
            ["name"] = check.Name,
            ["passed"] = check.Passed,
            ["severity"] = check.Severity,
            ["message"] = check.Message,
            ["details"] = check.Details,
        }).ToList();

        return new Dictionary<string, object?>
        {
            ["file_path"] = FilePath,
            ["total"] = Total,
            ["passed"] = Passed,
            ["warnings"] = Warnings,
            ["errors"] = Errors,
            ["checks"] = checks,
        };
    }
}
