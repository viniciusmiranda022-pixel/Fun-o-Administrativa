using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/[controller]")]
public class SettingsController : ControllerBase
{
    private readonly SettingsReader _reader;

    public SettingsController(SettingsReader reader)
    {
        _reader = reader;
    }

    [HttpGet]
    public ActionResult<ApiResponse<BackupSettings>> Get()
    {
        var s = _reader.Read();
        if (s == null)
        {
            return ApiResponse<BackupSettings>.Fail("settings.json não encontrado ou inválido");
        }
        return ApiResponse<BackupSettings>.Ok(s);
    }
}
