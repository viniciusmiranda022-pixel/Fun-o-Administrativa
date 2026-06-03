using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/unpacked-objects")]
public class UnpackedObjectsController : ControllerBase
{
    private readonly BackupReader _reader;

    public UnpackedObjectsController(BackupReader reader)
    {
        _reader = reader;
    }

    [HttpGet]
    public async Task<ActionResult<ApiResponse<UnpackedObjects>>> Get([FromQuery] string backupId)
    {
        if (string.IsNullOrWhiteSpace(backupId))
        {
            return ApiResponse<UnpackedObjects>.Fail("backupId is required");
        }

        var data = await _reader.GetUnpackedAsync(backupId);
        if (data == null)
        {
            return ApiResponse<UnpackedObjects>.Fail($"Backup {backupId} not found");
        }

        return ApiResponse<UnpackedObjects>.Ok(data);
    }
}
