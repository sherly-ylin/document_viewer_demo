using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;

namespace document_viewer_demo.Controllers
{
    public class SigningActionFilter: ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (filterContext.HttpContext.Request.Query["secureID"] != "123")
            {
                filterContext.Result = new Microsoft.AspNetCore.Mvc.ContentResult()
                {
                    Content = "Incorrect secureID: Access denied"
                };
            }

        }
    }
}