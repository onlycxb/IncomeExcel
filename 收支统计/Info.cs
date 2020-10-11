using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 收支统计
{
   public class Info
	{
		public string 会计凭证号
		{
			get;
			set;
		}

		public string 出纳凭证号
		{
			get;
			set;
		}

		public string 收款账户
		{
			get;
			set;
		}

		public DateTime 收款日期
		{
			get;
			set;
		}

		public string 发票号码
		{
			get;
			set;
		}

		public string 客户名称
		{
			get;
			set;
		}

		public DateTime 开票日期
		{
			get;
			set;
		}

		public string 商品名称
		{
			get;
			set;
		}

		public string 项目所属月份
		{
			get;
			set;
		}

		public string 摘要
		{
			get;
			set;
		}

		public decimal 金额
		{
			get;
			set;
		}

		public decimal 税额
		{
			get;
			set;
		}

		public decimal 价税合计
		{
			get;
			set;
		}

		public string 收款单号
		{
			get;
			internal set;
		}

		public string 发票管理区
		{
			get;
			internal set;
		}

		public string 实际管理区
		{
			get;
			internal set;
		}
	}
}
