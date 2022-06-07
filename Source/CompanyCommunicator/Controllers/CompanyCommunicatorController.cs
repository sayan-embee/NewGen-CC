
namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Configuration;

    /// <summary>
    /// Controller for the survey export data.
    /// </summary>
    [Route("api/companycommunicator")]
    [ApiController]
    public class CompanyCommunicatorController : ControllerBase
    {
        /// <summary>
        /// Get a sent notification by Id.
        /// </summary>
        /// <param name="id">Id of the requested sent notification.</param>
        /// <returns>Required sent notification.</returns>
        [HttpGet]
        [Route("tenantlist")]
        public async Task<IActionResult> GetSisterTenant()
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build();
            var sistertenant = configuration.GetSection("SisterTenantId").Value.ToString();
            return this.Ok(sistertenant);
        }
    }
}