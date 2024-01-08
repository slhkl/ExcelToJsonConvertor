using Business;
using Microsoft.AspNetCore.Mvc;

namespace Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        [HttpPost(nameof(ToJson))]
        public IActionResult ToJson(IFormFile formFile)
        {
            return Ok(formFile.OpenReadStream().ToJson());
        }
    }
}
