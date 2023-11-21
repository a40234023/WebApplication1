using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class ApiResult
    {
        /// <summary>
        /// 執行成功與否
        /// </summary>
        public string name { get; set; }
        /// <summary>
        /// 結果代碼(0000=成功，其餘為錯誤代號)
        /// </summary>
        public string code { get; set; }

    }
}