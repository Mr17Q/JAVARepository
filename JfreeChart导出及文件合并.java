package com.ipss.index.controller;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.awt.RenderingHints;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.chart.axis.CategoryAxis;
import org.jfree.chart.axis.CategoryLabelPositions;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.labels.ItemLabelAnchor;
import org.jfree.chart.labels.ItemLabelPosition;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.LineAndShapeRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.ui.TextAnchor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.support.RequestContextUtils;

import sun.misc.BASE64Encoder;

import com.ipss.index.domain.IndexBaseinfo;
import com.ipss.index.domain.IndexProfit;
import com.ipss.index.domain.IndexProfitGroup;
import com.ipss.index.domain.IndexZhcydl;
import com.ipss.index.domain.TBranchBenchmark;
import com.ipss.index.domain.TImprovementActions;
import com.ipss.index.domain.TImprovementMethod;
import com.ipss.index.domain.TImprovementUnit;
import com.ipss.index.domain.TIndexGdmh;
import com.ipss.index.domain.TIndexGdmhGrade;
import com.ipss.index.domain.TIndexGdmhGroup;
import com.ipss.index.domain.TIndexGdmhOptimal;
import com.ipss.index.domain.TQuestionMethod;
import com.ipss.index.dto.CwTmodeDTO;
import com.ipss.index.service.IIndexASService;
import com.ipss.index.service.IIndexBaseinfoService;
import com.ipss.index.service.IIndexCWService;
import com.ipss.index.service.IIndexClassService;
import com.ipss.index.service.IIndexExportService;
import com.ipss.index.service.IIndexFZBenchmarkYTDService;
import com.ipss.index.service.IIndexWZService;
import com.ipss.index.service.IIndexZhcydlService;
import com.ipss.index.service.ITImprovementActionsService;
import com.ipss.index.service.ITImprovementUnitService;
import com.ipss.index.service.ITIndexGdmhGradeService;
import com.ipss.index.service.ITIndexGdmhGroupService;
import com.ipss.index.service.ITIndexGdmhOptimalService;
import com.ipss.index.service.ITIndexGdmhService;
import com.ipss.index.service.ITQuestionMethodService;
import com.ipss.index.vo.FZBenchmarkYTDVO;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import exportentity.ReportDocExportEntity;
import exporttools.ChartUtils;
import exporttools.DocumentHandlerTwo;
import framework.common.dao.PaginationObject;
import framework.common.dao.PaginationParam;
import framework.common.tools.FrameworkConstants;
import framework.right.domain.Account;
import framework.right.domain.Organization;
import framework.right.service.IOrganMgrService;
import framework.right.utils.AccountHelper;
import framework.right.utils.UserView;
import framework.web.controller.BaseController;

/**
 * <br>
 * 类描述：
 * 
 * @ClassName: IndexExportController.java 指标导出控制器层
 * @author generate
 * @date 2016-4-19 9:45:16
 * @version V1.0
 */


/**
* Description：指标导出模块：各个部门word模板导出及整个对标报告生成
* @author：邹勤
* @date : 2017-05-20
* @version： V1.1 2017-05-26（V1.0 2017-05-20）
* @see ： 使用到了各个部门指标展示模块的数据获取方法，使用了ftl模板导出word，使用jacob合并word文件
*/  
@Controller
@RequestMapping("/web/index/indexExport")
public class IndexExportController extends BaseController {

	@Resource
	private IIndexBaseinfoService indexBaseinfoService;

	@Resource
	private IIndexClassService indexClassService;

	@Resource
	private IIndexZhcydlService indexZhcydlService;

	@Resource
	private ITQuestionMethodService tQuestionMethodService;

	@Resource
	private ITImprovementUnitService tImprovementUnitService;

	@Resource
	private ITImprovementActionsService tImprovementActionsService;

	@Resource
	private IIndexExportService indexExportService;

	@Resource
	private IIndexASService indexASService;

	@Resource
	private ITIndexGdmhService tIndexGdmhService;

	@Resource
	private ITIndexGdmhGroupService tIndexGdmhGroupService;

	@Resource
	private ITIndexGdmhGradeService tIndexGdmhGradeService;

	@Resource
	private ITIndexGdmhOptimalService tIndexGdmhOptimalService;

	@Resource
	private IIndexFZBenchmarkYTDService indexFZBenchmarkYTDService;

	@Resource
	private IOrganMgrService organMgrService;

	@Resource
	private IIndexWZService indexWZService;

	@Resource
	private IIndexCWService IndexCWService;

	// 当前单位背景色
	static String fillcolorNO = "auto";// 无色
	static String fillcolor = "CCCCFF";// 背景色

	/**
	 * 
	 * <br>
	 * 类描述：FZBenchmarkYTDVO排序实现类
	 * 
	 * @ClassName: MyComparator
	 * @author generate
	 * @date 2017-4-19 9:45:16
	 * @version V1.0
	 */
	private class MyComparator implements Comparator {

		public int compare(Object arg0, Object arg1) {
			FZBenchmarkYTDVO cop1 = (FZBenchmarkYTDVO) arg0;
			FZBenchmarkYTDVO cop2 = (FZBenchmarkYTDVO) arg1;
			return cop1.getOrderNoInt() - cop2.getOrderNoInt();
		}

	}

	/**
	 * 
	 * <br>
	 * 类描述：FZBenchmarkYTDVO排序实现类
	 * 
	 * @ClassName: MyComparator
	 * @author generate
	 * @date 2017-4-19 9:45:16
	 * @version V1.0
	 */
	private class MyComparator4JTNB implements Comparator {
		public int compare(Object arg0, Object arg1) {
			FZBenchmarkYTDVO cop1 = (FZBenchmarkYTDVO) arg0;
			FZBenchmarkYTDVO cop2 = (FZBenchmarkYTDVO) arg1;
			double temp = Double.valueOf(cop1.getIndexValue2()) - Double.valueOf(cop2.getIndexValue2());
			if (temp > 0){
				return 1;
			}else{
				return -1;
			}
		}
	}

	/**
	 * 进入对标导出
	 * 
	 * @param Model
	 * @return String
	 * @author zouqin
	 */
	@RequestMapping(value = "/toExportIndex")
	public final String toExportIndex(Model model) {
		logger.debug("\n toExportIndex() is running!");
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月");
		Calendar cal = Calendar.getInstance();
		cal.setTime(new Date());
		cal.add(Calendar.MONTH, -1);
		// 上个月
		String lastMonth = sdf.format(cal.getTime());
		model.addAttribute("lastMonth", lastMonth);
		model.addAttribute("orgId", this.getLoginAccount().getEmployee().getDept().getId());
		return "@export-index";
	}

	private String GetOrgId() {
		UserView uview = (UserView) AccountHelper.getAttribute(FrameworkConstants.USER_VIEW);
		return uview.getCurrentOrgId();
	}

	/**
	 * 计划部导出
	 * 
	 * @return ResponseEntity<byte[]> 返回导出对象
	 */
	@RequestMapping(value = "/toAll", method = RequestMethod.GET)
	@ResponseBody
	public ResponseEntity<byte[]> toAll(String date) throws Exception {
		logger.debug("\n toAll() is running!");
		SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy年MM月");
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		Map<String, Object> dataMap = new HashMap<String, Object>();
		DocumentHandlerTwo documentHandler = new DocumentHandlerTwo();
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;
		String month1 = sdf1.format(sdf2.parse(date));
		dataMap.put("month1", month1);
		// -----------------摘要数据----------------------
		UserView uview = (UserView) AccountHelper.getAttribute(FrameworkConstants.USER_VIEW);
		Account account = uview.getUser();
		String vOrgName = account.getEmployee().getComp().getName();
		String orgId = account.getEmployee().getComp().getId();
		// ----------------------------------表一/二开始-------------------------------------------
		List<IndexProfit> list1 = indexExportService.getOneTableData(date);
		dataMap.put("table1", list1);
		List<IndexProfit> list2 = indexExportService.getTwoTableData(date);
		dataMap.put("table2", list2);
		// ----------------------------------表一/二结束-------------------------------------------
		List<IndexProfit> list3 = indexExportService.getThreeTableData(date);
		for (IndexProfit index : list3) {
			if (null != index.getId() && !"".equals(index.getId())) {
				IndexBaseinfo indexBaseinfo = this.indexBaseinfoService.getById(index.getId());
				if (null != indexBaseinfo.getOrderMode() && "1".equals(indexBaseinfo.getOrderMode())) {// 指标值小为优
					if (Double.parseDouble(index.getYearAvgValue())
							- Double.parseDouble(index.getMonthAvgValue2()) > 0) {
						index.setYearAvgValue1("退步");
					} else if (Double.parseDouble(index.getYearAvgValue())
							- Double.parseDouble(index.getMonthAvgValue2()) == 0) {
						index.setYearAvgValue1("持平");
					} else {
						index.setYearAvgValue1("向好");
					}
				} else {
					if (Double.parseDouble(index.getYearAvgValue())
							- Double.parseDouble(index.getMonthAvgValue2()) > 0) {
						index.setYearAvgValue1("向好");
					} else if (Double.parseDouble(index.getYearAvgValue())
							- Double.parseDouble(index.getMonthAvgValue2()) == 0) {
						index.setYearAvgValue1("持平");
					} else {
						index.setYearAvgValue1("退步");
					}
				}
			} else {
				index.setYearAvgValue1("");
				if (index.getIndexName().equals("入厂标单与区域平均差值")) {
					if (Double.parseDouble(index.getYearAvgValue())
							- Double.parseDouble(index.getMonthAvgValue2()) > 0) {
						index.setYearAvgValue1("向好");
					} else if (Double.parseDouble(index.getYearAvgValue())
							- Double.parseDouble(index.getMonthAvgValue2()) == 0) {
						index.setYearAvgValue1("持平");
					} else {
						index.setYearAvgValue1("退步");
					}
				}
				if (index.getIndexName().equals("入厂、入炉标单差")) {
					if (Double.parseDouble(index.getYearAvgValue())
							- Double.parseDouble(index.getMonthAvgValue2()) > 0) {
						index.setYearAvgValue1("向好");
					} else if (Double.parseDouble(index.getYearAvgValue())
							- Double.parseDouble(index.getMonthAvgValue2()) == 0) {
						index.setYearAvgValue1("持平");
					} else {
						index.setYearAvgValue1("退步");
					}
				}
			}
		}
		dataMap.put("table3", list3);
		// ----------------------------------上月整改措施落实情况
		// 结束--------------------------------------------------
		String rootPath = RequestContextUtils.getWebApplicationContext(request).getServletContext().getRealPath("/");
		String strPath = rootPath + "model-head.doc";
		File file = new File(strPath);
		File fileParent = file.getParentFile();
		if (!fileParent.exists()) {
			fileParent.mkdirs();
		}
		try {
			file.createNewFile();
		} catch (IOException e) {
			e.printStackTrace();
		}
		File ff = documentHandler.createDoc(dataMap, "/conf/ftl", "model-head.ftl", strPath);
		HttpHeaders headers = new HttpHeaders();
		String exportName = "大唐华银电力股份有限公司";
		if (orgId.equals("40288f9b542d388801542d5099f90006") || orgId.equals("40288f9b542d388801542d51913d0008")
				|| orgId.equals("40288f9b542d388801542d51ea520009")
				|| orgId.equals("40288f9b542d388801542d50ea620007")) {
			exportName = vOrgName;
		}
		String filename = exportName + month1 + "份对标情况通报.doc";
		String fileName = new String(filename.getBytes("gb2312"), "iso-8859-1");// 为了解决中文名称乱码问题
		List list = new ArrayList();
		String file2 = rootPath + "model-dlyx.doc";
		String file3 = rootPath + "model-aqsc.doc";
		toExportJh("", date, "");
		toExportAs("", date, "");
		toWordExportRLWZ(date);
		toWordExportCW(date);
		String file4 = rootPath + "rlwz.doc";
		String file5 = rootPath + "cw.doc";
		String file1 = strPath;
		list.add(file1);
		list.add(file2);
		list.add(file3);
		list.add(file4);
		list.add(file5);
		File allfile = uniteDoc(list, rootPath +"model-all.doc");
		headers.setContentDispositionFormData("attachment", fileName);
		headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
		return new ResponseEntity<byte[]>(FileUtils.readFileToByteArray(allfile), headers, HttpStatus.OK);
	}

	/**
	 * 计划部导出
	 * 
	 * @return
	 */
	@RequestMapping(value = "/toExportJh", method = RequestMethod.GET)
	@ResponseBody
	public ResponseEntity<byte[]> toExportJh(String indexId, String date, String selId) throws Exception {
		indexId = "70f2e801-7d2f-4743-8390-3ca848ab5685";
		selId = "5d81245b-56ad-4f40-9250-64a392cab353";
		logger.debug("\n toExportJh() is running!");
		SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy年MM月");
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;
		IndexBaseinfo indexBaseinfo = this.indexBaseinfoService.getById(indexId);

		// -----------------摘要数据----------------------
		UserView uview = (UserView) AccountHelper.getAttribute(FrameworkConstants.USER_VIEW);
		Account account = uview.getUser();
		String orgId = account.getEmployee().getComp().getId();
		String orgName = account.getEmployee().getComp().getDescription();
		String vOrgName = account.getEmployee().getComp().getName();
		IndexProfit plan = indexClassService.getSummaryData(indexId, date, selId, orgId);
		List<String> list0 = new ArrayList<String>();
		if (null != plan) {
			list0.add(plan.getWorkPlan());
		} else {
			list0.add("");
		}
		// ----------------------------------表一开始-------------------------------------------
		List<IndexProfit> tablelist = new ArrayList<IndexProfit>();
		List<IndexProfit> tablelist1 = indexClassService.getMonthOneData(indexId, date, selId);
		Calendar cal = Calendar.getInstance();
		cal.setTime(sdf2.parse(date));
		cal.add(Calendar.MONTH, -1);
		// 上个月
		String lastMonth = sdf2.format(cal.getTime());
		String month1 = sdf1.format(sdf2.parse(date));
		String month2 = sdf1.format(cal.getTime()) + "份";
		boolean flag = false;
		List<IndexProfit> tablelist2 = indexClassService.getMonthOneData(indexId, lastMonth, selId);
		if (null != tablelist2 && tablelist2.size() > 0) {
			if (null != tablelist1 && tablelist1.size() > 0) {
				for (int i = 0; i < tablelist1.size(); i++) {
					IndexProfit pro = tablelist1.get(i);
					// 添加背景色标志
					if (pro.getSysId().equals(orgId)) {
						pro.setIsCurrOrg("1");
						flag = true;
					} else {
						pro.setIsCurrOrg("0");
					}
					for (int j = 0; j < tablelist2.size(); j++) {
						if (pro.getSysId().equals(tablelist2.get(j).getSysId())) {
							pro.setMonthAvgValue1(tablelist2.get(j).getAchieveValue());
							tablelist.add(pro);
						}
					}
				}
			}
		} else {
			tablelist = tablelist1;
		}
		// 默认给公司加背景色
		if (!flag) {
			tablelist.get(0).setIsCurrOrg("1");
		}
		Calendar cal2 = Calendar.getInstance();
		cal2.setTime(sdf2.parse(date));
		cal2.add(Calendar.YEAR, -1);
		// 上年同月
		String lastYear = sdf2.format(cal2.getTime());
		List<IndexProfit> tablelist3 = indexClassService.getMonthOneData(indexId, lastYear, selId);
		if (null != tablelist && tablelist.size() > 0) {
			for (int i = 0; i < tablelist.size(); i++) {
				IndexProfit pro = tablelist.get(i);
				for (int j = 0; j < tablelist3.size(); j++) {
					if (pro.getSysId().equals(tablelist3.get(j).getSysId())) {
						pro.setYearAvgValue1(tablelist3.get(j).getAddupValue());
					}
				}
			}
		}

		DocumentHandlerTwo documentHandler = new DocumentHandlerTwo();
		Map<String, Object> dataMap = new HashMap<String, Object>();
		dataMap.put("month1", month1);
		dataMap.put("month2", date.split("-")[0] + "年1-" + Integer.parseInt(date.split("-")[1]) + "月");
		// 组装数据
		ReportDocExportEntity reportDocExportEntity = new ReportDocExportEntity();
		reportDocExportEntity.setTitleLevelOne("大唐华银电力股份有限公司" + month1 + "对标情况通报");
		DecimalFormat df = (DecimalFormat) NumberFormat.getInstance();
		df.setMaximumFractionDigits(2);
		// 设置每一行的数据
		for (int i = 0; i < tablelist.size(); i++) {
			String monthval = df.format(Double.parseDouble(tablelist.get(i).getAchieveValue())
					- Double.parseDouble(tablelist.get(i).getMonthAvgValue1()));
			String yearval = df.format(Double.parseDouble(tablelist.get(i).getAddupValue())
					- Double.parseDouble(tablelist.get(i).getYearAvgValue1()));
			tablelist.get(i).setMonthAvgValue(monthval);
			tablelist.get(i).setYearAvgValue(yearval);
		}
		dataMap.put("table1list", tablelist);
		// ----------------------------------表一结束---------------------------------------------------
		// ----------------------------------表二开始---------------------------------------------------
		List<IndexProfitGroup> table2list = indexClassService.getMonthTwoData(indexId, date, selId, "2");
		List<IndexProfitGroup> lasttable2list = indexClassService.getMonthTwoData(indexId, lastMonth, selId, "2");
		List<IndexProfitGroup> tableT2list = indexClassService.getMonthTwoData(indexId, date, selId, "1");
		List<IndexProfit> listtable2 = new ArrayList<IndexProfit>();
		List<IndexProfit> table2 = new ArrayList<IndexProfit>();
		if (null != table2list && table2list.size() > 0 && null != lasttable2list && lasttable2list.size() > 0
				&& null != tableT2list && tableT2list.size() > 0) {
			IndexProfit pro = new IndexProfit();
			pro.setOrgName("区域平均");
			pro.setMonthAvgValue(table2list.get(0).getMonthAvgValue());
			pro.setYearAvgValue(tableT2list.get(0).getYearAvgValue());
			pro.setMonthAvgValue1("");
			pro.setMonthAvgValue2("");
			pro.setYearAvgValue1("");
			pro.setYearAvgValue2("");
			table2.add(pro);

			IndexProfit pro1 = new IndexProfit();
			pro1.setOrgName("华能");
			pro1.setMonthAvgValue(table2list.get(0).getHnGroup());
			pro1.setMonthAvgValue1(table2list.get(0).getHnGroupOrder());
			int cval = Integer.parseInt(table2list.get(0).getHnGroupOrder())
					- Integer.parseInt(lasttable2list.get(0).getHnGroupOrder());
			if (cval == 0) {
				pro1.setMonthAvgValue2("-");
			} else if (cval > 0) {
				pro1.setMonthAvgValue2("↓" + cval);
			} else {
				pro1.setMonthAvgValue2("↑" + (cval + "").split("-")[1]);
			}
			pro1.setYearAvgValue(tableT2list.get(0).getHnGroup());
			pro1.setYearAvgValue1(df.format(
					Double.parseDouble(tableT2list.get(0).getHnGroup()) - Double.parseDouble(pro.getYearAvgValue())));
			pro1.setYearAvgValue2(tableT2list.get(0).getHnGroupOrder());

			IndexProfit pro2 = new IndexProfit();
			pro2.setOrgName("大唐");
			pro2.setMonthAvgValue(table2list.get(0).getDtGroup());
			pro2.setMonthAvgValue1(table2list.get(0).getDtGroupOrder());
			cval = Integer.parseInt(table2list.get(0).getDtGroupOrder())
					- Integer.parseInt(lasttable2list.get(0).getDtGroupOrder());
			if (cval == 0) {
				pro2.setMonthAvgValue2("-");
			} else if (cval > 0) {
				pro2.setMonthAvgValue2("↓" + cval);
			} else {
				pro2.setMonthAvgValue2("↑" + (cval + "").split("-")[1]);
			}
			pro2.setYearAvgValue(tableT2list.get(0).getDtGroup());
			pro2.setYearAvgValue1(df.format(
					Double.parseDouble(tableT2list.get(0).getDtGroup()) - Double.parseDouble(pro.getYearAvgValue())));
			pro2.setYearAvgValue2(tableT2list.get(0).getDtGroupOrder());

			IndexProfit pro3 = new IndexProfit();
			pro3.setOrgName("长安");
			pro3.setMonthAvgValue(table2list.get(0).getCaGroup());
			pro3.setMonthAvgValue1(table2list.get(0).getCaGroupOrder());
			cval = Integer.parseInt(table2list.get(0).getCaGroupOrder())
					- Integer.parseInt(lasttable2list.get(0).getCaGroupOrder());
			if (cval == 0) {
				pro3.setMonthAvgValue2("-");
			} else if (cval > 0) {
				pro3.setMonthAvgValue2("↓" + cval);
			} else {
				pro3.setMonthAvgValue2("↑" + (cval + "").split("-")[1]);
			}
			pro3.setYearAvgValue(tableT2list.get(0).getCaGroup());
			pro3.setYearAvgValue1(df.format(
					Double.parseDouble(tableT2list.get(0).getCaGroup()) - Double.parseDouble(pro.getYearAvgValue())));
			pro3.setYearAvgValue2(tableT2list.get(0).getCaGroupOrder());

			IndexProfit pro4 = new IndexProfit();
			pro4.setOrgName("华电");
			pro4.setMonthAvgValue(table2list.get(0).getHdGroup());
			pro4.setMonthAvgValue1(table2list.get(0).getHdGroupOrder());
			cval = Integer.parseInt(table2list.get(0).getHdGroupOrder())
					- Integer.parseInt(lasttable2list.get(0).getHdGroupOrder());
			if (cval == 0) {
				pro4.setMonthAvgValue2("-");
			} else if (cval > 0) {
				pro4.setMonthAvgValue2("↓" + cval);
			} else {
				pro4.setMonthAvgValue2("↑" + (cval + "").split("-")[1]);
			}
			pro4.setYearAvgValue(tableT2list.get(0).getHdGroup());
			pro4.setYearAvgValue1(df.format(
					Double.parseDouble(tableT2list.get(0).getHdGroup()) - Double.parseDouble(pro.getYearAvgValue())));
			pro4.setYearAvgValue2(tableT2list.get(0).getHdGroupOrder());

			IndexProfit pro5 = new IndexProfit();
			pro5.setOrgName("国电");
			pro5.setMonthAvgValue(table2list.get(0).getGdGroup());
			pro5.setMonthAvgValue1(table2list.get(0).getGdGroupOrder());
			cval = Integer.parseInt(table2list.get(0).getGdGroupOrder())
					- Integer.parseInt(lasttable2list.get(0).getGdGroupOrder());
			if (cval == 0) {
				pro5.setMonthAvgValue2("-");
			} else if (cval > 0) {
				pro5.setMonthAvgValue2("↓" + cval);
			} else {
				pro5.setMonthAvgValue2("↑" + (cval + "").split("-")[1]);
			}
			pro5.setYearAvgValue(tableT2list.get(0).getGdGroup());
			pro5.setYearAvgValue1(df.format(
					Double.parseDouble(tableT2list.get(0).getGdGroup()) - Double.parseDouble(pro.getYearAvgValue())));
			pro5.setYearAvgValue2(tableT2list.get(0).getGdGroupOrder());

			IndexProfit pro6 = new IndexProfit();
			pro6.setOrgName("国电投");
			pro6.setMonthAvgValue(table2list.get(0).getTzGroup());
			pro6.setMonthAvgValue1(table2list.get(0).getTzGroupOrder());
			cval = Integer.parseInt(table2list.get(0).getTzGroupOrder())
					- Integer.parseInt(lasttable2list.get(0).getTzGroupOrder());
			if (cval == 0) {
				pro6.setMonthAvgValue2("-");
			} else if (cval > 0) {
				pro6.setMonthAvgValue2("↓" + cval);
			} else {
				pro6.setMonthAvgValue2("↑" + (cval + "").split("-")[1]);
			}
			pro6.setYearAvgValue(tableT2list.get(0).getTzGroup());
			pro6.setYearAvgValue1(df.format(
					Double.parseDouble(tableT2list.get(0).getTzGroup()) - Double.parseDouble(pro.getYearAvgValue())));
			pro6.setYearAvgValue2(tableT2list.get(0).getTzGroupOrder());

			IndexProfit pro7 = new IndexProfit();
			pro7.setOrgName("华润");
			pro7.setMonthAvgValue(table2list.get(0).getHrGroup());
			pro7.setMonthAvgValue1(table2list.get(0).getHrGroupOrder());
			cval = Integer.parseInt(table2list.get(0).getHrGroupOrder())
					- Integer.parseInt(lasttable2list.get(0).getHrGroupOrder());
			if (cval == 0) {
				pro7.setMonthAvgValue2("-");
			} else if (cval > 0) {
				pro7.setMonthAvgValue2("↓" + cval);
			} else {
				pro7.setMonthAvgValue2("↑" + (cval + "").split("-")[1]);
			}
			pro7.setYearAvgValue(tableT2list.get(0).getHrGroup());
			pro7.setYearAvgValue1(df.format(
					Double.parseDouble(tableT2list.get(0).getHrGroup()) - Double.parseDouble(pro.getYearAvgValue())));
			pro7.setYearAvgValue2(tableT2list.get(0).getHrGroupOrder());

			listtable2.add(pro1);
			listtable2.add(pro2);
			listtable2.add(pro3);
			listtable2.add(pro4);
			listtable2.add(pro5);
			listtable2.add(pro6);
			listtable2.add(pro7);
			// 排序
			Collections.sort(listtable2, new Comparator<IndexProfit>() {
				@Override
				public int compare(IndexProfit o1, IndexProfit o2) {
					int o1int = Integer.parseInt(o1.getYearAvgValue2());
					int o2int = Integer.parseInt(o2.getYearAvgValue2());
					return o1int - o2int;
				}
			});
			for (IndexProfit fit : listtable2) {
				table2.add(fit);
			}
		}
		dataMap.put("table2list", table2);
		// ----------------------------------表二结束---------------------------------------------------
		// ----------------------------------图一开始---------------------------------------------------
		DefaultCategoryDataset dataset = new DefaultCategoryDataset();
		// 图片数据
		String[] colors = new String[table2.size()];
		for (int i = 0; i < table2.size(); i++) {
			IndexProfit index = table2.get(i);
			if (index.getOrgName().equals("大唐")) {
				colors[i] = "#C0504D";// 红色
			} else {
				colors[i] = "#4F81BD";// 蓝色
			}
			dataset.addValue(Double.parseDouble(index.getYearAvgValue()), "", index.getOrgName());
		}
		dataMap.put("map1",
				createBarChart(dataset, "", "",
						"1-" + Integer.parseInt(date.split("-")[1]) + "月份公司" + indexBaseinfo.getIndexName() + "对标",
						"barGroup.png", false, colors));
		// ----------------------------------图一结束---------------------------------------------------
		// ----------------------------------表三开始---------------------------------------------------
		// 排序
		List<IndexProfit> tablelist4 = indexClassService.getMonthThreeData(indexId, date, selId);
		List<IndexProfit> copytablelist = new ArrayList<IndexProfit>();
		List<IndexProfit> newtablelist = new ArrayList<IndexProfit>();
		// 为空排最后
		newtablelist.add(tablelist4.get(0));
		for (int i = 1; i < tablelist4.size(); i++) {
			IndexProfit pros = tablelist4.get(i);
			if (null == pros.getYearSortNO() || "".equals(pros.getYearSortNO())) {
				copytablelist.add(pros);
			} else {
				newtablelist.add(pros);
			}
		}
		Collections.sort(newtablelist, new Comparator<IndexProfit>() {
			@Override
			public int compare(IndexProfit o1, IndexProfit o2) {
				int o1int = Integer.parseInt(o1.getYearSortNO());
				int o2int = Integer.parseInt(o2.getYearSortNO());
				return o1int - o2int;
			}
		});
		newtablelist.addAll(copytablelist);
		// 上个月
		List<IndexProfit> tablelist5 = indexClassService.getMonthThreeData(indexId, lastMonth, selId);
		List<IndexProfit> copytablelist2 = new ArrayList<IndexProfit>();
		List<IndexProfit> newtablelist2 = new ArrayList<IndexProfit>();
		// 为空排最后
		for (IndexProfit pros : tablelist5) {
			if (null == pros.getYearSortNO() || "".equals(pros.getYearSortNO())) {
				copytablelist2.add(pros);
			} else {
				newtablelist2.add(pros);
			}
		}
		// 排序
		Collections.sort(newtablelist2, new Comparator<IndexProfit>() {
			@Override
			public int compare(IndexProfit o1, IndexProfit o2) {
				int o1int = Integer.parseInt(o1.getYearSortNO());
				int o2int = Integer.parseInt(o2.getYearSortNO());
				return o1int - o2int;
			}
		});
		newtablelist2.addAll(copytablelist2);

		// 添加背景色标志
		boolean flags = false;
		for (int i = 0; i < newtablelist.size(); i++) {
			IndexProfit proi = newtablelist.get(i);
			if (null != proi.getSysId() && !"".equals(proi.getSysId()) && proi.getSysId().equals(orgId)) {
				proi.setIsCurrOrg("1");
				flags = true;
			} else {
				proi.setIsCurrOrg("0");
			}
		}
		// 默认给 株洲/湘潭/耒阳/金竹山
		if (!flags) {
			for (int i = 0; i < newtablelist.size(); i++) {
				IndexProfit prop = tablelist4.get(i);
				if (null != prop.getSysId() && !"".equals(prop.getSysId())) {
					if (prop.getSysId().equals("40288f9b542d388801542d5099f90006")
							|| prop.getSysId().equals("40288f9b542d388801542d51913d0008")
							|| prop.getSysId().equals("40288f9b542d388801542d51ea520009")
							|| prop.getSysId().equals("40288f9b542d388801542d50ea620007")) {
						prop.setIsCurrOrg("1");
					} else {
						prop.setIsCurrOrg("0");
					}
				} else {
					prop.setIsCurrOrg("0");
				}
			}
		}

		// 区域平均值比
		String yearValue = newtablelist.get(0).getYearAvgValue();

		// 添加箭头值
		for (int z = 1; z < newtablelist.size(); z++) {
			IndexProfit in = newtablelist.get(z);
			// 设置与区域平均值比
			String yearval = df.format(Double.parseDouble(in.getAddupValue()) - Double.parseDouble(yearValue));
			in.setYearAvgValue(yearval);
			if (null != newtablelist2 && newtablelist2.size() > 0) {
				for (IndexProfit newin : newtablelist2) {
					if (in.getOrgName().equals(newin.getOrgName())) {
						int val = Integer.parseInt(in.getMonthSortNO()) - Integer.parseInt(newin.getMonthSortNO());
						if (val == 0) {
							in.setMonthAvgValue("-");
						} else if (val > 0) {
							in.setMonthAvgValue("↓" + val);
						} else {
							in.setMonthAvgValue("↑" + (val + "").split("-")[1]);
						}
					}
				}
			}
		}

		// 设置第一条区域平均值
		newtablelist.get(0).setAchieveValue(newtablelist.get(0).getMonthAvgValue());
		newtablelist.get(0).setAddupValue(newtablelist.get(0).getYearAvgValue());
		newtablelist.get(0).setMonthSortNO("");
		newtablelist.get(0).setYearSortNO("");
		newtablelist.get(0).setIsCurrOrg("0");
		dataMap.put("table3list", newtablelist);
		// ----------------------------------表三结束---------------------------------------------------
		// ----------------------------------图二开始---------------------------------------------------
		DefaultCategoryDataset dataset2 = new DefaultCategoryDataset();
		String[] colors2 = new String[newtablelist.size()];
		for (int i = 0; i < newtablelist.size(); i++) {
			IndexProfit index = newtablelist.get(i);
			if (null != index.getIsCurrOrg() && index.getIsCurrOrg().equals("1")) {
				colors2[i] = "#C0504D";// 红色
			} else {
				colors2[i] = "#4F81BD";// 蓝色
			}
			dataset2.addValue(Double.parseDouble(index.getAddupValue()), "", index.getOrgName());
		}
		dataMap.put("map2", createBarChart(dataset2, "", "", "各单位本年度" + indexBaseinfo.getIndexName() + "对标",
				"barGroup2.png", false, colors2));
		// ----------------------------------图二结束---------------------------------------------------
		// ----------------------------------综合厂用电率----------------------------------------------
		// ----------------------------------表四开始---------------------------------------------------
		// ------------------表格数据---------------------
		List<IndexZhcydl> zhtablelist = new ArrayList<IndexZhcydl>();
		List<IndexZhcydl> zhtablelist1 = indexZhcydlService.getDatas("607491f3-e950-431b-8afa-32502ba2a7fd", date);
		// 上个月
		List<IndexZhcydl> zhtablelist2 = indexZhcydlService.getDatas("607491f3-e950-431b-8afa-32502ba2a7fd", lastMonth);
		if (null != zhtablelist2 && zhtablelist2.size() > 0) {
			if (null != zhtablelist1 && zhtablelist1.size() > 0) {
				for (int i = 0; i < zhtablelist1.size(); i++) {
					IndexZhcydl index = zhtablelist1.get(i);
					for (int j = 0; j < zhtablelist2.size(); j++) {
						IndexZhcydl lastindex = zhtablelist2.get(j);
						if (index.getOrgName().equals(lastindex.getOrgName())) {
							if (null != lastindex.getZmonthValue() && null != lastindex.getFmonthValue()) {
								String lastValue = df.format(Double.parseDouble(lastindex.getZmonthValue())
										- Double.parseDouble(lastindex.getFmonthValue()));
								index.setLastValue(lastValue);
							} else {
								index.setLastValue("");
							}
							zhtablelist.add(index);
						}
					}
				}
			}
		} else {
			zhtablelist = zhtablelist1;
		}
		List<IndexZhcydl> table4 = new ArrayList<IndexZhcydl>();
		List<IndexZhcydl> table5 = new ArrayList<IndexZhcydl>();
		List<IndexZhcydl> table6 = new ArrayList<IndexZhcydl>();
		List<IndexZhcydl> table7 = new ArrayList<IndexZhcydl>();// 用于图排序与表格排序不同
		List<IndexZhcydl> table5temp = new ArrayList<IndexZhcydl>();
		List<IndexZhcydl> table6temp = new ArrayList<IndexZhcydl>();
		double firstvalue = 0;
		for (int i = 0; i < zhtablelist.size(); i++) {
			IndexZhcydl zh = zhtablelist.get(i);
			if (null == zh.getZmonthValue() || "".equals(zh.getZmonthValue())) {
				zh.setZmonthValue("0");
			}
			if (null == zh.getFmonthValue() || "".equals(zh.getFmonthValue())) {
				zh.setFmonthValue("0");
			}
			if (null == zh.getLastValue() || "".equals(zh.getLastValue())) {
				zh.setLastValue("0");
			}
			if (null == zh.getZyearValue() || "".equals(zh.getZyearValue())) {
				zh.setZyearValue("0");
			}
			if (null == zh.getFyearValue() || "".equals(zh.getFyearValue())) {
				zh.setFyearValue("0");
			}
			// 表1
			if (zh.getOrgType().equals("0") || zh.getOrgType().equals("1")) {
				if (i == 0) {
					zh.setValue1(df
							.format(Double.parseDouble(zh.getZmonthValue()) - Double.parseDouble(zh.getFmonthValue())));
					zh.setValue2("");
					zh.setValue3(df.format(Double.parseDouble(zh.getZmonthValue())
							- Double.parseDouble(zh.getFmonthValue()) - Double.parseDouble(zh.getLastValue())));
					firstvalue = Double.parseDouble(zh.getZmonthValue()) - Double.parseDouble(zh.getFmonthValue());
				} else {
					zh.setValue1(df
							.format(Double.parseDouble(zh.getZmonthValue()) - Double.parseDouble(zh.getFmonthValue())));
					zh.setValue2(df.format(Double.parseDouble(zh.getZmonthValue())
							- Double.parseDouble(zh.getFmonthValue()) - firstvalue));
					zh.setValue3(df.format(Double.parseDouble(zh.getZmonthValue())
							- Double.parseDouble(zh.getFmonthValue()) - Double.parseDouble(zh.getLastValue())));
				}
				table4.add(zh);
			}
		}
		dataMap.put("table4list", table4);
		// ----------------------------------表四结束---------------------------------------------------
		// ----------------------------------表五开始---------------------------------------------------

		for (int i = 0; i < zhtablelist1.size(); i++) {
			IndexZhcydl zh = zhtablelist1.get(i);
			IndexZhcydl zz = new IndexZhcydl();
			zz.setOrgName(zh.getOrgName());
			zz.setOrgType(zh.getOrgType());
			zz.setZmonthValue(zh.getZmonthValue());
			zz.setFmonthValue(zh.getFmonthValue());
			zz.setZyearValue(zh.getZyearValue());
			zz.setFyearValue(zh.getFyearValue());
			zz.setValue1(zh.getValue1());
			zz.setValue2(zh.getValue2());
			zz.setValue3(zh.getValue3());
			// 表2、3
			if (zz.getOrgType().equals("0")) {
				zz.setValue1(
						df.format(Double.parseDouble(zz.getZyearValue()) - Double.parseDouble(zz.getFyearValue())));
				zz.setValue2("");
				table5.add(zz);
				table6.add(zz);
				table7.add(zz);
			}
			if (zz.getOrgType().equals("2")) {
				zz.setValue1(
						df.format(Double.parseDouble(zz.getZyearValue()) - Double.parseDouble(zz.getFyearValue())));
				table5temp.add(zz);
			}
			if (zz.getOrgType().equals("1") || zz.getOrgType().equals("3")) {
				zz.setValue1(
						df.format(Double.parseDouble(zz.getZyearValue()) - Double.parseDouble(zz.getFyearValue())));
				table6temp.add(zz);
			}
		}
		// 排序
		Collections.sort(table5temp, new Comparator<IndexZhcydl>() {
			@Override
			public int compare(IndexZhcydl o1, IndexZhcydl o2) {
				Double o1int = Double.parseDouble(o1.getValue1());
				Double o2int = Double.parseDouble(o2.getValue1());
				return o2int.compareTo(o1int);
			}
		});
		for (int i = 0; i < table5temp.size(); i++) {
			table5temp.get(i).setValue2(i + 1 + "");
		}
		table5.addAll(table5temp);
		dataMap.put("table5list", table5);
		// ----------------------------------表五结束---------------------------------------------------
		// ----------------------------------图三开始---------------------------------------------------
		DefaultCategoryDataset dataset3 = new DefaultCategoryDataset();
		String[] colors3 = new String[table5.size()];
		for (int i = 0; i < table5.size(); i++) {
			IndexZhcydl index = table5.get(i);
			if (index.getOrgName().equals("大唐")) {
				colors3[i] = "#C0504D";// 红色
			} else {
				colors3[i] = "#4F81BD";// 蓝色
			}
			dataset3.addValue(Double.parseDouble(index.getZyearValue()), "", index.getOrgName());
		}
		dataMap.put("map3", createBarChart(dataset3, "", "",
				"1-" + Integer.parseInt(date.split("-")[1]) + "月份各发电集团" + indexBaseinfo.getIndexName() + "对比（%）",
				"barGroup3.png", false, colors3));
		// ----------------------------------图三结束---------------------------------------------------
		// ----------------------------------表六开始---------------------------------------------------
		// 排序
		Collections.sort(table6temp, new Comparator<IndexZhcydl>() {
			@Override
			public int compare(IndexZhcydl o1, IndexZhcydl o2) {
				Double o1int = Double.parseDouble(o1.getValue1());
				Double o2int = Double.parseDouble(o2.getValue1());
				return o2int.compareTo(o1int);
			}
		});
		// 背景颜色
		boolean flagss = false;
		for (int i = 0; i < table6temp.size(); i++) {
			IndexZhcydl proi = table6temp.get(i);
			if (null != proi.getOrgName() && !"".equals(proi.getOrgName()) && proi.getOrgName().equals(orgName)) {
				proi.setValue3("1");
				flagss = true;
			} else {
				proi.setValue3("0");
			}
			proi.setValue2(i + 1 + "");
		}
		if (!flagss) {
			for (int i = 0; i < table6temp.size(); i++) {
				if (table6temp.get(i).getOrgName().equals("株洲") || table6temp.get(i).getOrgName().equals("耒阳")
						|| table6temp.get(i).getOrgName().equals("金竹山")
						|| table6temp.get(i).getOrgName().equals("湘潭")) {
					table6temp.get(i).setValue3("1");
				} else {
					table6temp.get(i).setValue3("0");
				}
			}
		}
		table6.addAll(table6temp);
		dataMap.put("table6list", table6);
		// ----------------------------------表六结束---------------------------------------------------
		// ----------------------------------图四开始---------------------------------------------------
		// 排序
		Collections.sort(table6temp, new Comparator<IndexZhcydl>() {
			@Override
			public int compare(IndexZhcydl o1, IndexZhcydl o2) {
				Double o1int = Double.parseDouble(o1.getZyearValue());
				Double o2int = Double.parseDouble(o2.getZyearValue());
				return o2int.compareTo(o1int);
			}
		});
		table7.addAll(table6temp);
		DefaultCategoryDataset dataset4 = new DefaultCategoryDataset();
		String[] colors4 = new String[table7.size()];
		for (int i = 0; i < table7.size(); i++) {
			IndexZhcydl index = table7.get(i);
			if (index.getValue3().equals("1")) {
				colors4[i] = "#C0504D";// 红色
			} else {
				colors4[i] = "#4F81BD";// 蓝色
			}
			dataset4.addValue(Double.parseDouble(index.getZyearValue()), "", index.getOrgName());
		}
		dataMap.put("map4", createBarChart(dataset4, "", "",
				"1-" + Integer.parseInt(date.split("-")[1]) + "月份各统调发电厂" + indexBaseinfo.getIndexName() + "对比（%）",
				"barGroup4.png", false, colors4));
		// ----------------------------------图四结束--------------------------------------------------
		// ----------------------------------问题措施开始--------------------------------------------------
		Map<String, Object> conditions = new HashMap<String, Object>();
		conditions.put("dataPeriod", date.replace("-", ""));
		conditions.put("catalog", "dlyx");
		String qxtype = uview.getEmployee().getDept().getType();
		if ("3".equals(qxtype)) { // 选择部门查询 。计划部
			conditions.put("deptname", uview.getOrgName());
		}
		conditions.put("qxtype", qxtype); // 机构类型
		this.getPaginationParam().setPageNum(10000);
		PaginationObject<TQuestionMethod> tmpPaginObject = this.tQuestionMethodService
				.getPaginationObjectByParams4HQL(conditions, this.getPaginationParam());
		List<TQuestionMethod> question = new ArrayList<TQuestionMethod>();
		for (int i = 0; i < tmpPaginObject.getResultList().size(); i++) {
			TQuestionMethod qu = tmpPaginObject.getResultList().get(i);
			TQuestionMethod tQuestionMethod = (TQuestionMethod) this.tQuestionMethodService.getById(qu.getId());
			Map<String, Object> conditions2 = new HashMap<String, Object>();
			conditions2.put("id", qu.getId());
			PaginationObject<TImprovementMethod> tmpPaginObject2 = this.tQuestionMethodService
					.SearchTImprovementMethod(conditions2, this.getPaginationParam());
			tQuestionMethod.settImprovementMethod(tmpPaginObject2.getResultList());
			// 添加文字前的数字
			for (int j = 0; j < tQuestionMethod.gettImprovementMethod().size(); j++) {
				TImprovementMethod method = tQuestionMethod.gettImprovementMethod().get(j);
				method.setMethod(j + 1 + "." + method.getMethod() + "（责任单位：" + method.getUnitName() + "；完成时间："
						+ method.getCompleteTime() + "）");
			}
			tQuestionMethod.setQuestion("（" + ToCH(i + 1) + "）" + tQuestionMethod.getQuestion());
			// tQuestionMethod.setReason(tQuestionMethod.getReason().replace("\r\n",
			// "<w:br/>"));
			String[] strs = tQuestionMethod.getReason().split("\r\n");
			if (strs.length > 0) {
				String ss = "";
				for (int t = 0; t < strs.length; t++) {
					if (!strs[t].equals("")) {
						ss += strs[t]
								+ "</w:t></w:r></w:p><w:p wsp:rsidR=\"00C20E64\" wsp:rsidRPr=\"00A44351\" wsp:rsidRDefault=\"000D44CF\" wsp:rsidP=\"00C20E64\"><w:pPr><w:spacing w:line=\"560\" w:line-rule=\"exact\"/><w:ind w:first-line-chars=\"200\" w:first-line=\"640\"/><w:rPr><w:rFonts w:ascii=\"仿宋_GB2312\" w:fareast=\"仿宋_GB2312\"/><wx:font wx:val=\"仿宋_GB2312\"/><w:sz w:val=\"32\"/><w:sz-cs w:val=\"32\"/></w:rPr></w:pPr><w:r wsp:rsidRPr=\"00A44351\"><w:rPr><w:rFonts w:ascii=\"仿宋_GB2312\" w:fareast=\"仿宋_GB2312\" w:hint=\"fareast\"/><wx:font wx:val=\"仿宋_GB2312\"/><w:sz w:val=\"32\"/><w:sz-cs w:val=\"32\"/></w:rPr><w:t>";
					}
				}
				tQuestionMethod.setReason(ss);
			}
			question.add(tQuestionMethod);
		}
		dataMap.put("question", question);
		// ----------------------------------问题措施结束--------------------------------------------------
		// ----------------------------------上月整改措施落实情况
		// 开始--------------------------------------------------
		List<TImprovementUnit> act = new ArrayList<TImprovementUnit>();
		Map<String, Object> conditions4 = new HashMap<String, Object>();
		// conditions4.put("orgId",
		// this.getLoginAccount().getEmployee().getDept().getId());
		conditions4.put("orgId", "");
		conditions4.put("catalog", "dlyx");
		conditions4.put("dataPeriod", lastMonth.replace("-", ""));
		conditions4.put("dutyOrgId", "");
		conditions4.put("questionOrgId", account.getEmployee().getComp().getId());
		PaginationObject<TImprovementUnit> tmpPaginObject4 = this.tImprovementUnitService
				.searchtImprovementUnit(conditions4, this.getPaginationParam());
		List<TImprovementUnit> fiTrachargedetail = tmpPaginObject4.getResultList();
		// 中文数字并拼接新对象
		String tempstr = "";
		int x = 1;
		int y = 1;
		for (int i = 0; i < fiTrachargedetail.size(); i++) {
			TImprovementUnit tu = fiTrachargedetail.get(i);
			if (!tu.getQuestion().equals(tempstr)) {
				y = 1;
				tempstr = tu.getQuestion();
				tu.setQuestion("（" + ToCH(x) + "）关于“" + tu.getQuestion() + "”整改措施落实情况：");
				Map<String, Object> conditions5 = new HashMap<String, Object>();
				conditions5.put("improvementUnitid", tu.getId());
				PaginationObject<TImprovementActions> tmpPaginObject5 = this.tImprovementActionsService
						.searchttImprovementActions(conditions5, this.getPaginationParam());
				List<TImprovementActions> actions = tmpPaginObject5.getResultList();
				for (int j = 0; j < actions.size(); j++) {
					TImprovementActions action = actions.get(j);
					action.setActions(y + "." + action.getActions().split("。:")[0].split("。（")[0].split("（责")[0] + "。");
					y++;
				}
				tu.getProActions().addAll(actions);
				act.add(tu);
				x++;
			} else {
				for (TImprovementUnit unit : act) {
					if (unit.getQuestion().equals("（" + ToCH(x - 1) + "）关于“" + tu.getQuestion() + "”整改措施落实情况：")) {
						Map<String, Object> conditions5 = new HashMap<String, Object>();
						conditions5.put("improvementUnitid", tu.getId());
						PaginationObject<TImprovementActions> tmpPaginObject5 = this.tImprovementActionsService
								.searchttImprovementActions(conditions5, this.getPaginationParam());
						List<TImprovementActions> actions = tmpPaginObject5.getResultList();
						for (int j = 0; j < actions.size(); j++) {
							TImprovementActions action = actions.get(j);
							action.setActions(
									y + "." + action.getActions().split("。:")[0].split("。（")[0].split("（责")[0] + "。");
							y++;
						}
						unit.getProActions().addAll(actions);
					}
				}
			}
		}
		dataMap.put("act", act);
		// ----------------------------------上月整改措施落实情况
		// 结束--------------------------------------------------
		String rootPath = RequestContextUtils.getWebApplicationContext(request).getServletContext().getRealPath("/");
		String strPath = rootPath + "model-dlyx.doc";
		File file = new File(strPath);
		File fileParent = file.getParentFile();
		if (!fileParent.exists()) {
			fileParent.mkdirs();
		}
		try {
			file.createNewFile();
		} catch (IOException e) {
			e.printStackTrace();
		}
		File ff = documentHandler.createDoc(dataMap, "/conf/ftl", "model_jh.ftl", strPath);
		HttpHeaders headers = new HttpHeaders();
		String exportName = "大唐华银电力股份有限公司";
		if (!!flags) {
			exportName = vOrgName;
		}
		String filename = exportName + month1 + "份电量营销对标情况通报.doc";
		String fileName = new String(filename.getBytes("gb2312"), "iso-8859-1");// 为了解决中文名称乱码问题
		headers.setContentDispositionFormData("attachment", fileName);
		headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
		return new ResponseEntity<byte[]>(FileUtils.readFileToByteArray(ff), headers, HttpStatus.OK);
	}

	/**
	 * 安生部导出
	 * 
	 * @return
	 */
	@RequestMapping(value = "/toExportAs", method = RequestMethod.GET)
	@ResponseBody
	public ResponseEntity<byte[]> toExportAs(String indexId, String date, String selId) throws Exception {
		indexId = "d2e163ae-c215-44f5-bb6f-87494a1a46ed";
		selId = "5d81245b-56ad-4f40-9250-64a392cab353";
		logger.debug("\n toExportJh() is running!");
		SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy年MM月");
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;
		IndexBaseinfo indexBaseinfo = this.indexBaseinfoService.getById(indexId);

		// -----------------摘要数据----------------------
		UserView uview = (UserView) AccountHelper.getAttribute(FrameworkConstants.USER_VIEW);
		Account account = uview.getUser();
		String orgId = account.getEmployee().getComp().getId();
		String vOrgName = account.getEmployee().getComp().getName();
		IndexProfit plan = indexClassService.getSummaryData(indexId, date, selId, orgId);
		List<String> list0 = new ArrayList<String>();
		if (null != plan) {
			list0.add(plan.getWorkPlan());
		} else {
			list0.add("");
		}
		DocumentHandlerTwo documentHandler = new DocumentHandlerTwo();
		Map<String, Object> dataMap = new HashMap<String, Object>();
		// ----------------------------------表一开始-------------------------------------------
		List<IndexProfit> tablelist = new ArrayList<IndexProfit>();
		List<IndexProfit> tablelist1 = indexClassService.getMonthOneData(indexId, date, selId);
		Calendar cal = Calendar.getInstance();
		cal.setTime(sdf2.parse(date));
		cal.add(Calendar.MONTH, -1);
		// 上个月
		String lastMonth = sdf2.format(cal.getTime());
		String month1 = sdf1.format(sdf2.parse(date));
		dataMap.put("month1", month1);
		dataMap.put("month2", date.split("-")[0] + "年1-" + Integer.parseInt(date.split("-")[1]) + "月");

		Map<String, Object> conditions = new HashMap<String, Object>();
		if (StringUtils.isNotEmpty(indexId)) {
			conditions.put("indexBaseinfoId", indexId);
		} else {
			conditions.put("indexBaseinfoId", "");
		}
		if (StringUtils.isNotEmpty(date)) {
			conditions.put("effectCycle", date);
		} else {
			conditions.put("effectCycle", "");
		}
		PaginationParam paginationParam = this.getPaginationParam();
		paginationParam.setPageNum(-1);
		paginationParam.setPageSize(200);
		PaginationObject<TIndexGdmh> tmpPaginObject = this.tIndexGdmhService.getPaginationObjectByParams4HQL(conditions,
				paginationParam);
		List<TIndexGdmh> table1 = tmpPaginObject.getResultList();
		if (null != table1 && table1.size() > 0) {
			for (TIndexGdmh in : table1) {
				if (null == in.getOrgName() || "".equals(in.getOrgName())) {
					table1 = new ArrayList<TIndexGdmh>();
				}
			}
		} else {
			table1 = new ArrayList<TIndexGdmh>();
		}
		dataMap.put("table1", table1);
		// ----------------------------------表一结束-------------------------------------------
		// ----------------------------------表二开始-------------------------------------------
		PaginationObject<TIndexGdmhGroup> tmpPaginObject2 = this.tIndexGdmhGroupService
				.getPaginationObjectByParams4HQL(conditions, paginationParam);
		List<TIndexGdmhGroup> table2 = tmpPaginObject2.getResultList();
		if (null != table2 && table2.size() > 0) {
			for (TIndexGdmhGroup in : table2) {
				if (null == in.getIndexBaseinfoName() || "".equals(in.getIndexBaseinfoName())) {
					table2 = new ArrayList<TIndexGdmhGroup>();
				}
			}
		} else {
			table2 = new ArrayList<TIndexGdmhGroup>();
		}
		dataMap.put("table2", table2);
		// ----------------------------------表二结束-------------------------------------------
		// ----------------------------------表三开始-------------------------------------------
		PaginationObject<TIndexGdmhGrade> tmpPaginObject3 = this.tIndexGdmhGradeService
				.getPaginationObjectByParams4HQL(conditions, paginationParam);
		List<TIndexGdmhGrade> table3 = tmpPaginObject3.getResultList();
		dataMap.put("table3", table3);
		// ----------------------------------表三结束-------------------------------------------
		// ----------------------------------表四开始-------------------------------------------
		PaginationObject<TIndexGdmhOptimal> tmpPaginObject4 = this.tIndexGdmhOptimalService
				.getPaginationObjectByParams4HQL(conditions, paginationParam);
		List<TIndexGdmhOptimal> table4 = tmpPaginObject4.getResultList();
		dataMap.put("table4", table4);
		// ----------------------------------表四结束-------------------------------------------

		// ----------------------------------配煤参烧开始-------------------------------------------
		indexId = "da8011d6-1d13-4451-85df-81129fd106cd";
		Organization org = this.organMgrService.getOrganizationById(uview.getCurrentOrgId());
		String curOrgId = uview.getCurrentOrgId();
		String fzOrgId = "40288d904815e0eb014815ff5f590008";
		String jcOrgIds = "'40288f9b542d388801542d50ea620007','40288f9b542d388801542d51ea520009','40288f9b542d388801542d5099f90006','40288f9b542d388801542d51913d0008'";
		List<FZBenchmarkYTDVO> maplist = new ArrayList<FZBenchmarkYTDVO>();
		List<IndexProfit> tablelistFZ = indexFZBenchmarkYTDService.getIndexProfitData(indexId, date);
		List<IndexProfit> tablelistJC = indexFZBenchmarkYTDService.getIndexProfitDataByOrgIds(indexId, date, jcOrgIds,
				" addup_value desc ");

		// 插入分公司数据
		java.text.DecimalFormat df = new java.text.DecimalFormat("#.00"); // 小数位数格式化两位
		if (tablelistFZ.size() > 0) {
			// 分公司
			String isColor = "";
			if (fzOrgId.equals(curOrgId)) {
				isColor = "1";
			}
			maplist.add(new FZBenchmarkYTDVO("公司", tablelistFZ.get(0).getYearAvgValue(),
					tablelistFZ.get(0).getAchieveValue(), tablelistFZ.get(0).getAddupValue(), "", "-", -1, isColor));

		}
		if (tablelistJC.size() > 0) {
			// 循环将基层单位数据加到tablelist
			for (int i = 0; i < tablelistJC.size(); i++) {
				double indexVlaue = 0.0;
				double indexVlaue1 = 0.0;
				double indexVlaue2 = 0.0;
				String isColor = "";
				if (tablelistJC.get(i).getYearAvgValue() != null && tablelistJC.get(i).getYearAvgValue() != "") {
					if (tablelistJC.get(i).getSysId().equals(curOrgId)) {
						isColor = "1";
					}
				}
				if (tablelistJC.get(i).getYearAvgValue() != null && tablelistJC.get(i).getYearAvgValue() != "") {
					indexVlaue = Double.valueOf(tablelistJC.get(i).getYearAvgValue());
				}
				if (tablelistJC.get(i).getAchieveValue() != null && tablelistJC.get(i).getAchieveValue() != "") {
					indexVlaue1 = Double.valueOf(tablelistJC.get(i).getAchieveValue());
				}
				if (tablelistJC.get(i).getAddupValue() != null && tablelistJC.get(i).getAddupValue() != "") {
					indexVlaue2 = Double.valueOf(tablelistJC.get(i).getAddupValue());
				}
				String orgDes = this.organMgrService.getOrganizationById(tablelistJC.get(i).getSysId())
						.getDescription();
				maplist.add(new FZBenchmarkYTDVO(orgDes, String.valueOf(indexVlaue), String.valueOf(indexVlaue1),
						String.valueOf(indexVlaue2), "", String.valueOf(i + 1), i + 1, isColor));
			}
		}
		DefaultCategoryDataset dataset = new DefaultCategoryDataset();
		// 图片数据
		for (int i = 0; i < maplist.size(); i++) {
			FZBenchmarkYTDVO index = maplist.get(i);
			dataset.addValue(null == index.getIndexValue() ? 0 : Double.parseDouble(index.getIndexValue()), "年计划",
					index.getOrgName());
			dataset.addValue(null == index.getIndexValue1() ? 0 : Double.parseDouble(index.getIndexValue1()), "当月",
					index.getOrgName());
			dataset.addValue(null == index.getIndexValue2() ? 0 : Double.parseDouble(index.getIndexValue2()), "年累计",
					index.getOrgName());
		}
		dataMap.put("map1", createBarChart(dataset, "", "", "烧效益标单贡献与年度计划对比图", "barGroup.png", true, null));

		// ---------map2---------------------
		DefaultCategoryDataset dataset2 = new DefaultCategoryDataset();
		String newyear = Integer.parseInt(date.split("-")[0]) - 1 + "";
		int nowmonth = Integer.parseInt(date.split("-")[1]);
		int nowyear = Integer.parseInt(date.split("-")[0]);
		List<IndexProfit> maplist1 = new ArrayList<IndexProfit>();
		List<IndexProfit> maplist2 = new ArrayList<IndexProfit>();
		// 去年
		for (int i = 1; i <= 12; i++) {
			String newdate = newyear + "-" + fillZero(i);
			IndexProfit pro = indexClassService.getMonthFourData(indexId, newdate, selId, orgId);
			pro.setOrgName(monthToUppder(i));
			maplist1.add(pro);
		}
		// 今年
		for (int j = 1; j <= nowmonth; j++) {
			String newdate = nowyear + "-" + fillZero(j);
			IndexProfit pro = indexClassService.getMonthFourData(indexId, newdate, selId, orgId);
			pro.setOrgName(monthToUppder(j));
			maplist2.add(pro);
		}
		for (int i = 0; i < maplist1.size(); i++) {
			IndexProfit index1 = maplist1.get(i);
			IndexProfit index2 = null;
			if (maplist2.size() > i) {
				index2 = maplist2.get(i);
			}
			dataset2.addValue(null == index1.getAddupValue() ? 0 : Double.parseDouble(index1.getAddupValue()), "2016年",
					index1.getOrgName());
			if (null != index2)
				dataset2.addValue(null == index2.getAddupValue() ? 0 : Double.parseDouble(index2.getAddupValue()),
						"2017年", index2.getOrgName());
		}
		dataMap.put("map2", createLineChart(dataset2, "", "", "华银公司月累计掺烧效益标单贡献趋势对比图", "barGroup.png", true, null));

		// ----------------------------------配煤参烧结束-------------------------------------------
		// ----------------------------------问题措施开始--------------------------------------------------
		Map<String, Object> qconditions = new HashMap<String, Object>();
		qconditions.put("dataPeriod", date.replace("-", ""));
		qconditions.put("catalog", "aqsc");
		String qxtype = uview.getEmployee().getDept().getType();
		if ("3".equals(qxtype)) { // 选择部门查询 。安生部
			qconditions.put("deptname", uview.getOrgName());
		}
		qconditions.put("qxtype", qxtype); // 机构类型
		this.getPaginationParam().setPageNum(10000);
		PaginationObject<TQuestionMethod> qtmpPaginObject = this.tQuestionMethodService
				.getPaginationObjectByParams4HQL(qconditions, this.getPaginationParam());
		List<TQuestionMethod> question = new ArrayList<TQuestionMethod>();
		for (int i = 0; i < qtmpPaginObject.getResultList().size(); i++) {
			TQuestionMethod qu = qtmpPaginObject.getResultList().get(i);
			TQuestionMethod tQuestionMethod = (TQuestionMethod) this.tQuestionMethodService.getById(qu.getId());
			Map<String, Object> conditions2 = new HashMap<String, Object>();
			conditions2.put("id", qu.getId());
			PaginationObject<TImprovementMethod> qtmpPaginObject2 = this.tQuestionMethodService
					.SearchTImprovementMethod(conditions2, this.getPaginationParam());
			tQuestionMethod.settImprovementMethod(qtmpPaginObject2.getResultList());
			// 添加文字前的数字
			for (int j = 0; j < tQuestionMethod.gettImprovementMethod().size(); j++) {
				TImprovementMethod method = tQuestionMethod.gettImprovementMethod().get(j);
				method.setMethod(j + 1 + "." + method.getMethod() + "（责任单位：" + method.getUnitName() + "；完成时间："
						+ method.getCompleteTime() + "）");
			}
			tQuestionMethod.setQuestion("（" + ToCH(i + 1) + "）" + tQuestionMethod.getQuestion());
			// tQuestionMethod.setReason(tQuestionMethod.getReason().replace("\r\n",
			// "<w:br/>"));
			String[] strs = tQuestionMethod.getReason().split("\r\n");
			if (strs.length > 0) {
				String ss = "";
				for (int t = 0; t < strs.length; t++) {
					if (!strs[t].equals("")) {
						ss += strs[t]
								+ "</w:t></w:r></w:p><w:p wsp:rsidR=\"00C20E64\" wsp:rsidRPr=\"00A44351\" wsp:rsidRDefault=\"000D44CF\" wsp:rsidP=\"00C20E64\"><w:pPr><w:spacing w:line=\"560\" w:line-rule=\"exact\"/><w:ind w:first-line-chars=\"200\" w:first-line=\"640\"/><w:rPr><w:rFonts w:ascii=\"仿宋_GB2312\" w:fareast=\"仿宋_GB2312\"/><wx:font wx:val=\"仿宋_GB2312\"/><w:sz w:val=\"32\"/><w:sz-cs w:val=\"32\"/></w:rPr></w:pPr><w:r wsp:rsidRPr=\"00A44351\"><w:rPr><w:rFonts w:ascii=\"仿宋_GB2312\" w:fareast=\"仿宋_GB2312\" w:hint=\"fareast\"/><wx:font wx:val=\"仿宋_GB2312\"/><w:sz w:val=\"32\"/><w:sz-cs w:val=\"32\"/></w:rPr><w:t>";
					}
				}
				tQuestionMethod.setReason(ss);
			}
			question.add(tQuestionMethod);
		}
		dataMap.put("question", question);
		// ----------------------------------问题措施结束--------------------------------------------------
		// ----------------------------------上月整改措施落实情况
		// 开始--------------------------------------------------
		List<TImprovementUnit> act = new ArrayList<TImprovementUnit>();
		Map<String, Object> conditions4 = new HashMap<String, Object>();
		// conditions4.put("orgId",
		// this.getLoginAccount().getEmployee().getDept().getId());
		conditions4.put("orgId", "");
		conditions4.put("catalog", "aqsc");
		conditions4.put("dataPeriod", lastMonth.replace("-", ""));
		conditions4.put("dutyOrgId", "");
		conditions4.put("questionOrgId", account.getEmployee().getComp().getId());
		PaginationObject<TImprovementUnit> atmpPaginObject4 = this.tImprovementUnitService
				.searchtImprovementUnit(conditions4, this.getPaginationParam());
		List<TImprovementUnit> fiTrachargedetail = atmpPaginObject4.getResultList();
		// 中文数字并拼接新对象
		String tempstr = "";
		int x = 1;
		int y = 1;
		for (int i = 0; i < fiTrachargedetail.size(); i++) {
			TImprovementUnit tu = fiTrachargedetail.get(i);
			if (!tu.getQuestion().equals(tempstr)) {
				y = 1;
				tempstr = tu.getQuestion();
				tu.setQuestion("（" + ToCH(x) + "）关于“" + tu.getQuestion() + "”整改措施落实情况：");
				Map<String, Object> conditions5 = new HashMap<String, Object>();
				conditions5.put("improvementUnitid", tu.getId());
				PaginationObject<TImprovementActions> tmpPaginObject5 = this.tImprovementActionsService
						.searchttImprovementActions(conditions5, this.getPaginationParam());
				List<TImprovementActions> actions = tmpPaginObject5.getResultList();
				for (int j = 0; j < actions.size(); j++) {
					TImprovementActions action = actions.get(j);
					action.setActions(y + "." + action.getActions().split("。:")[0] + "。");
					y++;
				}
				tu.getProActions().addAll(actions);
				act.add(tu);
				x++;
			} else {
				for (TImprovementUnit unit : act) {
					if (unit.getQuestion().equals("（" + ToCH(x - 1) + "）关于“" + tu.getQuestion() + "”整改措施落实情况：")) {
						Map<String, Object> conditions5 = new HashMap<String, Object>();
						conditions5.put("improvementUnitid", tu.getId());
						PaginationObject<TImprovementActions> tmpPaginObject5 = this.tImprovementActionsService
								.searchttImprovementActions(conditions5, this.getPaginationParam());
						List<TImprovementActions> actions = tmpPaginObject5.getResultList();
						for (int j = 0; j < actions.size(); j++) {
							TImprovementActions action = actions.get(j);
							action.setActions(y + "." + action.getActions().split("。:")[0] + "。");
							y++;
						}
						unit.getProActions().addAll(actions);
					}
				}
			}
		}
		dataMap.put("act", act);
		// ----------------------------------上月整改措施落实情况
		// 结束--------------------------------------------------
		String rootPath = RequestContextUtils.getWebApplicationContext(request).getServletContext().getRealPath("/");
		String strPath = rootPath + "model-aqsc.doc";
		File file = new File(strPath);
		File fileParent = file.getParentFile();
		if (!fileParent.exists()) {
			fileParent.mkdirs();
		}
		try {
			file.createNewFile();
		} catch (IOException e) {
			e.printStackTrace();
		}
		File ff = documentHandler.createDoc(dataMap, "/conf/ftl", "model_as.ftl", strPath);
		HttpHeaders headers = new HttpHeaders();
		String exportName = "大唐华银电力股份有限公司";
		if (orgId.equals("40288f9b542d388801542d5099f90006") || orgId.equals("40288f9b542d388801542d51913d0008")
				|| orgId.equals("40288f9b542d388801542d51ea520009")
				|| orgId.equals("40288f9b542d388801542d50ea620007")) {
			exportName = vOrgName;
		}
		String filename = exportName + month1 + "份安全生产对标情况通报.doc";
		String fileName = new String(filename.getBytes("gb2312"), "iso-8859-1");// 为了解决中文名称乱码问题
		headers.setContentDispositionFormData("attachment", fileName);
		headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
		return new ResponseEntity<byte[]>(FileUtils.readFileToByteArray(ff), headers, HttpStatus.OK);
	}

	/**
	 * 燃料物资部word导出
	 * 
	 * @return
	 */
	@RequestMapping(value = "/toWordExportRLWZ", method = RequestMethod.GET)
	@ResponseBody
	public ResponseEntity<byte[]> toWordExportRLWZ(String date) throws Exception {
		SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy年MM月");
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		String month1=sdf1.format(sdf2.parse(fdate));
		DocumentHandlerTwo documentHandler = new DocumentHandlerTwo();
		Map<String, Object> dataMap = new HashMap<String, Object>();
		UserView uview = (UserView) AccountHelper.getAttribute(FrameworkConstants.USER_VIEW);
		Organization org = this.organMgrService.getOrganizationById(uview.getCurrentOrgId());
		Account account = uview.getUser();
		String vOrgName = account.getEmployee().getComp().getName();
		String orgId = uview.getCurrentOrgId();
		String orgType = org.getType();
		String qxtype = uview.getEmployee().getDept().getType();
		String deptname = uview.getOrgName();
		String catalog = "rlwz";
		getData4RLWZ(dataMap, date, orgId, orgType, qxtype, deptname, catalog, "");
		String filePath = RequestContextUtils.getWebApplicationContext(request).getServletContext().getRealPath("/");
		;// 文档导出路径
		String strPath = filePath + "rlwz.doc";
		File file = new File(strPath);
		File fileParent = file.getParentFile();
		if (!fileParent.exists()) {
			fileParent.mkdirs();
		}
		try {
			file.createNewFile();
		} catch (IOException e) {
			e.printStackTrace();
		}
		File ff = documentHandler.createDoc(dataMap, "/conf/ftl", "rlwz.ftl", strPath);
		/*HttpHeaders headers = new HttpHeaders();
		String fileName = new String("rlwz.doc".getBytes("gb2312"), "iso-8859-1");// 为了解决中文乱码问题
		headers.setContentDispositionFormData("attachment", fileName);
		headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
		return new ResponseEntity<byte[]>(FileUtils.readFileToByteArray(ff), headers, HttpStatus.OK);*/
		
		 HttpHeaders headers = new HttpHeaders(); 
		 if(vOrgName.equals("大唐华银"))
	        {
	        	vOrgName = "大唐华银电力股份有限公司";
	        }
	        String filename=vOrgName+month1+"份燃料物资对标情况通报.doc";
	        String fileName=new String(filename.getBytes("gb2312"),"iso-8859-1");//为了解决中文乱码问题  
	        headers.setContentDispositionFormData("attachment", fileName);   
	        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);   
	        return new ResponseEntity<byte[]>(FileUtils.readFileToByteArray(ff),    
	                                          headers, HttpStatus.OK);
	}

	/**
	 * 财务部word导出
	 * 
	 * @return
	 */
	@RequestMapping(value = "/toWordExportCW", method = RequestMethod.GET)
	@ResponseBody
	public ResponseEntity<byte[]> toWordExportCW(String date) throws Exception {
		SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy年MM月");
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		String month1=sdf1.format(sdf2.parse(fdate));
		DocumentHandlerTwo documentHandler = new DocumentHandlerTwo();
		Map<String, Object> dataMap = new HashMap<String, Object>();
		UserView uview = (UserView) AccountHelper.getAttribute(FrameworkConstants.USER_VIEW);
		Organization org = this.organMgrService.getOrganizationById(uview.getCurrentOrgId());
		Account account = uview.getUser();
		String vOrgName = account.getEmployee().getComp().getName();
		String orgId = uview.getCurrentOrgId();
		String orgType = org.getType();
		String qxtype = uview.getEmployee().getDept().getType();
		String deptname = uview.getOrgName();
		String catalog = "cwjy";//
		getData4CW(dataMap, date, orgId, orgType, qxtype, deptname, catalog, "");
		String filePath = RequestContextUtils.getWebApplicationContext(request).getServletContext().getRealPath("/");
		;// 文档导出路径
		String strPath = filePath + "cw.doc";
		File file = new File(strPath);
		File fileParent = file.getParentFile();
		if (!fileParent.exists()) {
			fileParent.mkdirs();
		}
		try {
			file.createNewFile();
		} catch (IOException e) {
			e.printStackTrace();
		}
		File ff = documentHandler.createDoc(dataMap, "/conf/ftl", "cw.ftl", strPath);
		/*HttpHeaders headers = new HttpHeaders();
		String fileName = new String("cw.doc".getBytes("gb2312"), "iso-8859-1");// 为了解决中文名称乱码问题
		headers.setContentDispositionFormData("attachment", fileName);
		headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
		return new ResponseEntity<byte[]>(FileUtils.readFileToByteArray(ff), headers, HttpStatus.OK);*/
		
		 HttpHeaders headers = new HttpHeaders(); 
		 if(vOrgName.equals("大唐华银"))
	        {
	        	vOrgName = "大唐华银电力股份有限公司";
	        }
	        String filename=vOrgName+month1+"份财务经营对标情况通报.doc";
	        String fileName=new String(filename.getBytes("gb2312"),"iso-8859-1");//为了解决中文乱码问题  
	        headers.setContentDispositionFormData("attachment", fileName);   
	        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);   
	        return new ResponseEntity<byte[]>(FileUtils.readFileToByteArray(ff),    
	                                          headers, HttpStatus.OK);
	}

	/**
	 * 燃料物资部word导出 相关表格数据组装方法
	 *
	 */
	private void getData4RLWZ(Map dataMap, String date, String orgId, String orgType, String qxtype, String deptname,
			String catalog, String dutyOrgId) throws Exception {

		int nowYear = Integer.parseInt(date.split("-")[0]);
		int nowMonth = Integer.parseInt(date.split("-")[1]);
		String rcbdindexId = "80415ca3-6de7-4a9d-8505-0da4e168cde1";
		String rcbdZTQKselId = "761d3e63-a8ef-41c9-bd1a-d16cb5643e80";
		String rcbdQYDBselId = "b689f577-48c5-4016-b63e-b84d66ec8e44";
		String rcbdGHDDWDBselId = "b689f577-48c5-4016-b63e-b84d66ec8e55";
		String rcbdJTDBselId = "4a95680c-cdb4-4761-890a-728917145eec";
		String clbdcindexId = "7dd7c9dd-2fd0-4bc8-a965-8c47707cd565";
		String clbdcZTQKselId = "1b933333-67a1-4f87-8c1b-ce149de89894";
		String clbdcQYDBselId = "1b933333-67a1-4f87-8c1b-ce149de82383";
		String clbdcGHDDWDBselId = "1b933333-67a1-4f87-8c1b-ce149de82384";
		String jzcglindexId = "fe22eceb-efbb-4bae-a68f-755ac3df8c75";
		String jzcgHDselId = "";
		List<FZBenchmarkYTDVO> rcbdztqklists = getWZRLData4ZTQK(orgId, rcbdindexId, date, rcbdZTQKselId);
		List<FZBenchmarkYTDVO> rcbdqydblists = getWZRLData4WDJTDB(orgId, rcbdindexId, date, rcbdQYDBselId);
		List<FZBenchmarkYTDVO> rcbdghddwdblists = getWZRLData4QYDB(orgId, rcbdindexId, date, rcbdGHDDWDBselId, "2");
		List<FZBenchmarkYTDVO> rcbdjtdblists = getWZRLData4RCBDJTNBDB(orgId, rcbdindexId, date, rcbdJTDBselId);
		List<FZBenchmarkYTDVO> clbdcztqklists = getWZRLData4ZTQK(orgId, clbdcindexId, date, clbdcZTQKselId);
		List<FZBenchmarkYTDVO> clbdcqydblists = getWZRLData4WDJTDB(orgId, clbdcindexId, date, clbdcQYDBselId);
		List<FZBenchmarkYTDVO> clbdcghddwdblists = getWZRLData4QYDB(orgId, clbdcindexId, date, clbdcGHDDWDBselId, "1");
		Organization org = this.organMgrService.getOrganizationById(orgId);

		// 各火电单位标单差计划值完成表
		FZBenchmarkYTDVO clbdcghddwjhwc = new FZBenchmarkYTDVO("", "0", "0", "0", "0", "0", "", "", "", "", 0, "", "",
				"");
		if (null != clbdcztqklists && clbdcztqklists.size() > 0) {
			for (int i = 0; i < clbdcztqklists.size(); i++) {
				if (clbdcztqklists.get(i).getOrgName().equals("公司")||clbdcztqklists.get(i).getOrgName().equals("大唐")) {
					clbdcghddwjhwc.setIndexValue(clbdcztqklists.get(i).getIndexValue2());
				}
				if (clbdcztqklists.get(i).getOrgName().equals("湘潭")) {
					clbdcghddwjhwc.setIndexValue1(clbdcztqklists.get(i).getIndexValue2());
				}
				if (clbdcztqklists.get(i).getOrgName().equals("金竹山")) {
					clbdcghddwjhwc.setIndexValue2(clbdcztqklists.get(i).getIndexValue2());
				}
				if (clbdcztqklists.get(i).getOrgName().equals("耒阳")) {
					clbdcghddwjhwc.setIndexValue3(clbdcztqklists.get(i).getIndexValue2());
				}
				if (clbdcztqklists.get(i).getOrgName().equals("株洲")) {
					clbdcghddwjhwc.setIndexValue4(clbdcztqklists.get(i).getIndexValue2());
				}
			}
		}
		// 集中采购火电
		FZBenchmarkYTDVO jzcglhdM = new FZBenchmarkYTDVO("", "0", "0", "0", "0", "0", "", "", "", "", 0, "", "", "");
		FZBenchmarkYTDVO jzcglhdY = new FZBenchmarkYTDVO("", "0", "0", "0", "0", "0", "", "", "", "", 0, "", "", "");
		// 集中采购水电 张水 怀水 小洪 衡阳 新能源 欣正 先一
		FZBenchmarkYTDVO jzcglfhdM = new FZBenchmarkYTDVO("", "0", "0", "0", "0", "0", "0", "0", "0", "0", 0, "", "",
				"");
		FZBenchmarkYTDVO jzcglfhdY = new FZBenchmarkYTDVO("", "0", "0", "0", "0", "0", "0", "0", "0", "0", 0, "", "",
				"");

		List<IndexProfit> jzcglhdlistAvg = indexWZService.getTableDataAVG(jzcglindexId, date, jzcgHDselId);
		// 添加集团和公司值（水、火电一样）
		if (null != jzcglhdlistAvg && jzcglhdlistAvg.size() > 0) {
			if (null != jzcglhdlistAvg.get(0).getMonthAvgValue()) {
				jzcglhdM.setIndexValue(jzcglhdlistAvg.get(0).getMonthAvgValue());
				jzcglfhdM.setIndexValue(jzcglhdlistAvg.get(0).getMonthAvgValue());
			}
			if (null != jzcglhdlistAvg.get(0).getAchieveValue()) {
				jzcglhdM.setIndexValue1(jzcglhdlistAvg.get(0).getAchieveValue());
				jzcglfhdM.setIndexValue1(jzcglhdlistAvg.get(0).getAchieveValue());
			}
			if (null != jzcglhdlistAvg.get(0).getYearAvgValue()) {
				jzcglhdY.setIndexValue(jzcglhdlistAvg.get(0).getYearAvgValue());
				jzcglfhdY.setIndexValue(jzcglhdlistAvg.get(0).getYearAvgValue());
			}
			if (null != jzcglhdlistAvg.get(0).getAddupValue()) {
				jzcglhdY.setIndexValue1(jzcglhdlistAvg.get(0).getAddupValue());
				jzcglfhdY.setIndexValue1(jzcglhdlistAvg.get(0).getAddupValue());
			}
		}

		// 集中采购火电 月数据
		List<IndexProfit> jzcglhdlistM = indexWZService.getTableDataFIT(jzcglindexId, date, "1");// 月
		if (null != jzcglhdlistM && jzcglhdlistM.size() > 0) {
			for (int i = 0; i < jzcglhdlistM.size(); i++) {

				if (jzcglhdlistM.get(i).getOrgName().equals("金竹山")) {
					jzcglhdM.setIndexValue2(jzcglhdlistM.get(i).getAchieveValue());
				}
				if (jzcglhdlistM.get(i).getOrgName().equals("耒阳")) {
					jzcglhdM.setIndexValue3(jzcglhdlistM.get(i).getAchieveValue());
				}
				if (jzcglhdlistM.get(i).getOrgName().equals("湘潭")) {
					jzcglhdM.setIndexValue4(jzcglhdlistM.get(i).getAchieveValue());
				}
				if (jzcglhdlistM.get(i).getOrgName().equals("株洲")) {
					jzcglhdM.setIndexValue5(jzcglhdlistM.get(i).getAchieveValue());
				}
			}
		}
		// 集中采购火电 年数据
		List<IndexProfit> jzcglhdlistY = indexWZService.getTableDataFIT(jzcglindexId, date, "2");// 年
		if (null != jzcglhdlistY && jzcglhdlistY.size() > 0) {
			for (int i = 0; i < jzcglhdlistY.size(); i++) {

				if (jzcglhdlistY.get(i).getOrgName().equals("金竹山")) {
					jzcglhdY.setIndexValue2(jzcglhdlistY.get(i).getAddupValue());
				}
				if (jzcglhdlistY.get(i).getOrgName().equals("耒阳")) {
					jzcglhdY.setIndexValue3(jzcglhdlistY.get(i).getAddupValue());
				}
				if (jzcglhdlistY.get(i).getOrgName().equals("湘潭")) {
					jzcglhdY.setIndexValue4(jzcglhdlistY.get(i).getAddupValue());
				}
				if (jzcglhdlistY.get(i).getOrgName().equals("株洲")) {
					jzcglhdY.setIndexValue5(jzcglhdlistY.get(i).getAddupValue());
				}
			}
		}

		// 集中采购水电 月数据
		List<IndexProfit> jzcglfhdlistM = indexWZService.getTableDataOUTDC(jzcglindexId, date, "1");// 月
		if (null != jzcglfhdlistM && jzcglfhdlistM.size() > 0) {
			for (int i = 0; i < jzcglfhdlistM.size(); i++) {
				if (jzcglfhdlistM.get(i).getOrgName().equals("张水 ")) {
					jzcglfhdM.setIndexValue2(jzcglfhdlistM.get(i).getAchieveValue());
				}
				if (jzcglfhdlistM.get(i).getOrgName().equals("怀水")) {
					jzcglfhdM.setIndexValue3(jzcglfhdlistM.get(i).getAchieveValue());
				}
				if (jzcglfhdlistM.get(i).getOrgName().equals("小洪")) {
					jzcglfhdM.setIndexValue4(jzcglfhdlistM.get(i).getAchieveValue());
				}
				if (jzcglfhdlistM.get(i).getOrgName().equals("衡阳")) {
					jzcglfhdM.setIndexValue5(jzcglfhdlistM.get(i).getAchieveValue());
				}
				if (jzcglfhdlistM.get(i).getOrgName().equals("新能源")) {
					jzcglfhdM.setIndexValue6(jzcglfhdlistM.get(i).getAchieveValue());
				}
				if (jzcglfhdlistM.get(i).getOrgName().equals("欣正")) {
					jzcglfhdM.setAvgValue(jzcglfhdlistM.get(i).getAchieveValue());
				}
				if (jzcglfhdlistM.get(i).getOrgName().equals("先一")) {
					jzcglfhdM.setOrderNo(jzcglfhdlistM.get(i).getAchieveValue());
				}
			}
		}
		// 集中采购水电 年数据
		List<IndexProfit> jzcglfhdlistY = indexWZService.getTableDataOUTDC(jzcglindexId, date, "2");// 年
		if (null != jzcglfhdlistY && jzcglfhdlistY.size() > 0) {
			for (int i = 0; i < jzcglfhdlistY.size(); i++) {
				if (jzcglfhdlistY.get(i).getOrgName().equals("张水 ")) {
					jzcglfhdY.setIndexValue2(jzcglfhdlistY.get(i).getAddupValue());
				}
				if (jzcglfhdlistY.get(i).getOrgName().equals("怀水")) {
					jzcglfhdY.setIndexValue3(jzcglfhdlistY.get(i).getAddupValue());
				}
				if (jzcglfhdlistY.get(i).getOrgName().equals("小洪")) {
					jzcglfhdY.setIndexValue4(jzcglfhdlistY.get(i).getAddupValue());
				}
				if (jzcglfhdlistY.get(i).getOrgName().equals("衡阳")) {
					jzcglfhdY.setIndexValue5(jzcglfhdlistY.get(i).getAddupValue());
				}
				if (jzcglfhdlistY.get(i).getOrgName().equals("新能源")) {
					jzcglfhdY.setIndexValue6(jzcglfhdlistY.get(i).getAddupValue());
				}
				if (jzcglfhdlistY.get(i).getOrgName().equals("欣正")) {
					jzcglfhdY.setAvgValue(jzcglfhdlistY.get(i).getAddupValue());
				}
				if (jzcglfhdlistY.get(i).getOrgName().equals("先一")) {
					jzcglfhdY.setOrderNo(jzcglfhdlistY.get(i).getAddupValue());
				}
			}
		}
		dataMap.put("orgName", org.getName());
		dataMap.put("year", String.valueOf(nowYear));
		dataMap.put("month", nowMonth);
		// 入厂标单
		dataMap.put("rcbdztqkData", rcbdztqklists);
		dataMap.put("rcbdqydbData", rcbdqydblists);
		dataMap.put("rcbdghddwdbData", rcbdghddwdblists);
		dataMap.put("rcbdjtdbData", rcbdjtdblists);
		// 厂炉标单差
		dataMap.put("clbdcztqkData", clbdcztqklists);
		dataMap.put("clbdcqydbData", clbdcqydblists);
		dataMap.put("clbdcghddwdbData", clbdcghddwdblists);
		// 各火电单位标单差计划值完成表
		dataMap.put("clbdcghddwjhwcData", clbdcghddwjhwc);
		// 各火电单位集中采购率表
		dataMap.put("jzcglhdDataM", jzcglhdM);
		dataMap.put("jzcglhdDataY", jzcglhdY);
		// 各水电单位集中采购率表
		dataMap.put("jzcglfhdDataM", jzcglfhdM);
		dataMap.put("jzcglfhdDataY", jzcglfhdY);
		// 问题措施
		List<TQuestionMethod> question = getQuestionData(qxtype, deptname, catalog, date);
		dataMap.put("question", question);
		// 上月整改措施落实情况
		List<TImprovementUnit> act = getActData(orgId, catalog, date);
		dataMap.put("act", act);
	}

	// 获取总体情况表数据
	private List<FZBenchmarkYTDVO> getWZRLData4ZTQK(String orgId, String indexId, String date, String selId)
			throws Exception {
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;

		// ------------------表格数据---------------------
		List<FZBenchmarkYTDVO> tablelist = new ArrayList<FZBenchmarkYTDVO>();
		List<IndexProfit> tablelist1 = indexClassService.getMonthOneData(indexId, date, selId);
		Calendar cal = Calendar.getInstance();
		cal.setTime(sdf2.parse(date));
		cal.add(Calendar.YEAR, -1);
		// 上年同月
		String newDate = sdf2.format(cal.getTime());
		boolean flag = false;
		List<IndexProfit> tablelist3 = indexClassService.getMonthOneData(indexId, newDate, selId);
		// 环比
		Calendar ycal = Calendar.getInstance();
		ycal.setTime(sdf2.parse(date));
		ycal.add(Calendar.MONTH, -1);
		// 上个月
		String ynewDate = sdf2.format(ycal.getTime());
		List<IndexProfit> tablelist2 = indexClassService.getMonthOneData(indexId, ynewDate, selId);
		// 插入平均值
		java.text.DecimalFormat df = new java.text.DecimalFormat("#.##"); // 小数位数格式化两位
		if (null != tablelist1 && tablelist1.size() > 0) {
			for (int i = 0; i < tablelist1.size(); i++) {
				double indexVlaue = 0.0;
				double indexVlaue1 = 0.0;
				double indexVlaue2 = 0.0;
				double indexVlaue3 = 0.0;
				if (tablelist1.get(i).getAchieveValue() != null && tablelist1.get(i).getAchieveValue() != "") {
					indexVlaue = Double.valueOf(tablelist1.get(i).getAchieveValue());
				}
				if (tablelist1.get(i).getAddupValue() != null && tablelist1.get(i).getAddupValue() != "") {
					indexVlaue2 = Double.valueOf(tablelist1.get(i).getAddupValue());
				}
				if (null != tablelist2 && tablelist2.size() > 0) {
					for (int j = 0; j < tablelist2.size(); j++) {

						if (tablelist1.get(i).getSysId().equals(tablelist2.get(j).getSysId())) {
							if (tablelist2.get(j).getAchieveValue() != null
									&& tablelist2.get(j).getAchieveValue() != "") {
								indexVlaue1 = Double.valueOf(tablelist2.get(j).getAchieveValue());
							}
						}
					}
				}
				if (null != tablelist3 && tablelist3.size() > 0) {
					for (int k = 0; k < tablelist3.size(); k++) {

						if (tablelist1.get(i).getSysId().equals(tablelist3.get(k).getSysId())) {
							if (tablelist3.get(k).getAddupValue() != null && tablelist3.get(k).getAddupValue() != "") {
								indexVlaue3 = Double.valueOf(tablelist3.get(k).getAddupValue());
							}
						}
					}
				}
				if (tablelist1.get(i).getSysId().equals(orgId)) {
					tablelist.add(new FZBenchmarkYTDVO(tablelist1.get(i).getOrgName(),
							String.valueOf(df.format(indexVlaue)), String.valueOf(df.format(indexVlaue - indexVlaue1)),
							String.valueOf(df.format(indexVlaue2)),
							String.valueOf(df.format(indexVlaue2 - indexVlaue3)), "", "", "", "", "", 0, "", "",
							fillcolor));
				} else {
					tablelist.add(new FZBenchmarkYTDVO(tablelist1.get(i).getOrgName(),
							String.valueOf(df.format(indexVlaue)), String.valueOf(df.format(indexVlaue - indexVlaue1)),
							String.valueOf(df.format(indexVlaue2)),
							String.valueOf(df.format(indexVlaue2 - indexVlaue3)), "", "", "", "", "", 0, "", "",
							fillcolorNO));
				}
			}
		}
		return tablelist;
	}

	// 获取五大集团对标表数据
	private List<FZBenchmarkYTDVO> getWZRLData4WDJTDB(String orgId, String indexId, String date, String selId)
			throws Exception {
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;
		List<FZBenchmarkYTDVO> tablelist = new ArrayList<FZBenchmarkYTDVO>();
		// ------------------表格数据---------------------
		List<IndexProfit> avglist = indexFZBenchmarkYTDService.getIndexProfitData(indexId, date);
		List<CwTmodeDTO> tablelist1 = IndexCWService.getTableDataOUTFZ(indexId, date, selId);
		// 插入平均值
		java.text.DecimalFormat df = new java.text.DecimalFormat("#.##"); // 小数位数格式化两位
		double monthAvg = 0.0;
		double yearAvg = 0.0;
		if (avglist.size() > 0) {
			// 区域平均
			if (avglist.get(0).getFiveMonthAvgProfit() != null && avglist.get(0).getFiveMonthAvgProfit() != "") {
				monthAvg = Double.valueOf(avglist.get(0).getFiveMonthAvgProfit());
			}
			if (avglist.get(0).getFiveYearAvgProfit() != null && avglist.get(0).getFiveYearAvgProfit() != "") {
				yearAvg = Double.valueOf(avglist.get(0).getFiveYearAvgProfit());
			}
			// 集团
			tablelist.add(new FZBenchmarkYTDVO("平均值", String.valueOf(df.format(monthAvg)), "-", "-",
					String.valueOf(df.format(yearAvg)), "-", "-", "", "", "", -1, "", "", fillcolorNO));
		}
		for (int i = 0; i < tablelist1.size(); i++) {
			double monthVlaue = 0.0;
			double yearVlaue1 = 0.0;
			int sortNo = 0;
			String yueNo = "";
			String nianNo = "";
			if (tablelist1.get(i).getNianno() != null && tablelist1.get(i).getNianno() != "") {
				Integer it = new Integer(tablelist1.get(i).getNianno());
				sortNo = it.intValue();
			}
			if (tablelist1.get(i).getYueno() != null && tablelist1.get(i).getYueno() != "") {
				yueNo = tablelist1.get(i).getYueno();
			}
			if (tablelist1.get(i).getNianno() != null && tablelist1.get(i).getNianno() != "") {
				nianNo = tablelist1.get(i).getNianno();
			}
			if (tablelist1.get(i).getYuezhi() != null && tablelist1.get(i).getYuezhi() != "") {
				monthVlaue = Double.valueOf(tablelist1.get(i).getYuezhi());
			}
			if (tablelist1.get(i).getnianzhi() != null && tablelist1.get(i).getnianzhi() != "") {
				yearVlaue1 = Double.valueOf(tablelist1.get(i).getnianzhi());
			}
			if (!"国电投".equals(tablelist1.get(i).getCpname())) {
				if ("大唐".equals(tablelist1.get(i).getCpname())) {
					tablelist.add(new FZBenchmarkYTDVO(tablelist1.get(i).getCpname(),
							String.valueOf(df.format(monthVlaue)), String.valueOf(df.format(monthVlaue - monthAvg)),
							yueNo, String.valueOf(df.format(yearVlaue1)),
							String.valueOf(df.format(yearVlaue1 - yearAvg)), nianNo, "", "", "", sortNo, "", "",
							fillcolor));
				} else {
					tablelist.add(new FZBenchmarkYTDVO(tablelist1.get(i).getCpname(),
							String.valueOf(df.format(monthVlaue)), String.valueOf(df.format(monthVlaue - monthAvg)),
							yueNo, String.valueOf(df.format(yearVlaue1)),
							String.valueOf(df.format(yearVlaue1 - yearAvg)), nianNo, "", "", "", sortNo, "", "",
							fillcolorNO));
				}
			}
		}
		// 排序
		MyComparator myCompatator = new MyComparator();
		Collections.sort(tablelist, myCompatator);
		return tablelist;
	}

	// 获取区域对标表数据
	private List<FZBenchmarkYTDVO> getWZRLData4QYDB(String orgId, String indexId, String date, String selId,
			String avgFlag) throws Exception {
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;
		java.text.DecimalFormat df = new java.text.DecimalFormat("#.##"); // 小数位数格式化两位
		// ------------------表格数据---------------------
		List<FZBenchmarkYTDVO> tablelist = new ArrayList<FZBenchmarkYTDVO>();
		double monthAvg = 0.0;
		double yearAvg = 0.0;
		List<IndexProfit> avglist = indexFZBenchmarkYTDService.getIndexProfitData(indexId, date);

		List<IndexProfit> tablelist1 = IndexCWService.getTableData(indexId, date, selId);
		if (avglist.size() > 0) {
			
				// 区域平均
				if (avglist.get(0).getMonthAvgValue() != null && avglist.get(0).getMonthAvgValue() != "") {
					monthAvg = Double.valueOf(avglist.get(0).getMonthAvgValue());
				}
				if (avglist.get(0).getYearAvgValue() != null && avglist.get(0).getYearAvgValue() != "") {
					yearAvg = Double.valueOf(avglist.get(0).getYearAvgValue());
				}
			// 集团
			tablelist.add(new FZBenchmarkYTDVO("平均值", String.valueOf(df.format(monthAvg)), "-", "-",
					String.valueOf(df.format(yearAvg)), "-", "-", "", "", "", -1, "", "", fillcolorNO));
		}
		// 修改背景色
		if (orgId.equals("40288f9b542d388801542d5099f90006") || orgId.equals("40288f9b542d388801542d50ea620007")
				|| orgId.equals("40288f9b542d388801542d51ea520009")
				|| orgId.equals("40288f9b542d388801542d51913d0008")) {
			for (int i = 0; i < tablelist1.size(); i++) {
				if (!tablelist1.get(i).getSysId().equals(orgId)) {
					tablelist1.get(i).setAnnualObject("其他");
				}
			}
		}

		for (int i = 0; i < tablelist1.size(); i++) {
			double monthVlaue = 0.0;
			double yearVlaue1 = 0.0;
			int sortNo = 0;
			if (tablelist1.get(i).getYearSortNO() != null && tablelist1.get(i).getYearSortNO() != "") {
				Integer it = new Integer(tablelist1.get(i).getYearSortNO());
				sortNo = it.intValue();
			}
			if (tablelist1.get(i).getAchieveValue() != null && tablelist1.get(i).getAchieveValue() != "") {
				monthVlaue = Double.valueOf(tablelist1.get(i).getAchieveValue());
			}
			if (tablelist1.get(i).getAddupValue() != null && tablelist1.get(i).getAddupValue() != "") {
				yearVlaue1 = Double.valueOf(tablelist1.get(i).getAddupValue());
			}
			if ("大唐".equals(tablelist1.get(i).getAnnualObject())) {
				tablelist
						.add(new FZBenchmarkYTDVO(tablelist1.get(i).getOrgName(), String.valueOf(df.format(monthVlaue)),
								String.valueOf(df.format(monthVlaue - monthAvg)), tablelist1.get(i).getMonthSortNO(),
								String.valueOf(df.format(yearVlaue1)), String.valueOf(df.format(yearVlaue1 - yearAvg)),
								tablelist1.get(i).getYearSortNO(), "", "", "", sortNo, "", "", fillcolor));
			} else {
				tablelist
						.add(new FZBenchmarkYTDVO(tablelist1.get(i).getOrgName(), String.valueOf(df.format(monthVlaue)),
								String.valueOf(df.format(monthVlaue - monthAvg)), tablelist1.get(i).getMonthSortNO(),
								String.valueOf(df.format(yearVlaue1)), String.valueOf(df.format(yearVlaue1 - yearAvg)),
								tablelist1.get(i).getYearSortNO(), "", "", "", sortNo, "", "", fillcolorNO));
			}
		}
		return tablelist;
	}

	// 获取入厂标单集团内部对标表数据
	private List<FZBenchmarkYTDVO> getWZRLData4RCBDJTNBDB(String orgId, String indexId, String date, String selId)
			throws Exception {
		List<FZBenchmarkYTDVO> resultMap = new ArrayList<FZBenchmarkYTDVO>();
		List<FZBenchmarkYTDVO> tablelist = new ArrayList<FZBenchmarkYTDVO>();
		List<IndexProfit> tablelist1 = indexFZBenchmarkYTDService.getIndexProfitData(indexId, date);
		List<TBranchBenchmark> tablelist2 = indexFZBenchmarkYTDService.getBranchBenchmarkData(indexId, date, orgId,
				"Month_Sort_NO ");

		// 插入平均值
		java.text.DecimalFormat df = new java.text.DecimalFormat("#.##"); // 小数位数格式化两位

		double jtavg = 0.0;
		double jtIndexVlaue = 0.0;
		double avg = 0.0;
		double curIndexVlaue = 0.0;
		int sortNo = 0;
		if (tablelist1.size() > 0) {
			// 集团标单
			if (tablelist1.get(0).getYearAvgValue1() != null && tablelist1.get(0).getYearAvgValue1() != "") {
				jtIndexVlaue = Double.valueOf(tablelist1.get(0).getYearAvgValue1());
			}
			// 集团标单区域平均
			if (tablelist1.get(0).getYearAvgValue2() != null && tablelist1.get(0).getYearAvgValue2() != "") {
				jtavg = Double.valueOf(tablelist1.get(0).getYearAvgValue2());
			}
			// 集团
			resultMap.add(new FZBenchmarkYTDVO("集团", String.valueOf(df.format(jtIndexVlaue)),
					String.valueOf(df.format(jtavg)), String.valueOf(df.format(jtIndexVlaue - jtavg)), "-", "", "", "",
					"", "", -1, "", "", fillcolorNO));
			if (tablelist1.get(0).getAddupValue() != null && tablelist1.get(0).getAddupValue() != "") {
				curIndexVlaue = Double.valueOf(tablelist1.get(0).getAddupValue());
			}
			if (tablelist1.get(0).getYearAvgValue() != null && tablelist1.get(0).getYearAvgValue() != "") {
				avg = Double.valueOf(tablelist1.get(0).getYearAvgValue());
			}
			if (tablelist1.get(0).getYearSortNO() != null && tablelist1.get(0).getYearSortNO() != "") {
				Integer it = new Integer(tablelist1.get(0).getYearSortNO());
				sortNo = it.intValue();
			}
			// 湖南
			tablelist.add(new FZBenchmarkYTDVO("湖南", String.valueOf(df.format(curIndexVlaue)),
					String.valueOf(df.format(avg)), String.valueOf(df.format(curIndexVlaue - avg)),
					String.valueOf(sortNo), "", "", "", "", "", sortNo, "", "", fillcolor));
		}
		if (tablelist2.size() > 0) {
			// 循环将数据加到tablelist
			for (int i = 0; i < tablelist2.size(); i++) {
				double indexVlaue = 0.0;
				double avg1 = 0.0;
				int sortNo1 = 0;
				if (tablelist2.get(i).getYearSortNo() != null && tablelist2.get(i).getYearSortNo() != "") {
					Integer it = new Integer(tablelist2.get(i).getYearSortNo());
					sortNo1 = it.intValue();
				}
				if (tablelist2.get(i).getYearValue() != null && tablelist2.get(i).getYearValue() != "") {
					indexVlaue = Double.valueOf(tablelist2.get(i).getYearValue());
				}
				if (tablelist2.get(i).getSplitValue() != null && tablelist2.get(i).getSplitValue() != "") {
					avg1 = Double.valueOf(tablelist2.get(i).getSplitValue());
				}
				tablelist.add(
						new FZBenchmarkYTDVO(tablelist2.get(i).getBranchName(), String.valueOf(df.format(indexVlaue)),
								String.valueOf(df.format(avg1)), String.valueOf(df.format(indexVlaue - avg1)),
								String.valueOf(sortNo1), "", "", "", "", "", sortNo1, "", "", fillcolorNO));
			}
			// 排序
			MyComparator4JTNB myCompatator = new MyComparator4JTNB();
			Collections.sort(tablelist, myCompatator);

			if (tablelist.size() > 0) {
				// 循环将数据加到tablelist
				for (int i = 0; i < tablelist.size(); i++) {
					tablelist.get(i).setIndexValue3(String.valueOf(i + 1));
					resultMap.add(tablelist.get(i));
				}
			}
		}
		return resultMap;
	}

	// 获取问题措施数据
	private List<TQuestionMethod> getQuestionData(String qxtype, String deptname, String catalog, String date)
			throws Exception {
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;
		Map<String, Object> conditions = new HashMap<String, Object>();
		conditions.put("dataPeriod", date.replace("-", ""));
		conditions.put("catalog", catalog);
		if ("3".equals(qxtype)) { // 选择部门查询 。计划部
			conditions.put("deptname", deptname);
		}
		conditions.put("qxtype", qxtype); // 机构类型
		this.getPaginationParam().setPageNum(10000);
		PaginationObject<TQuestionMethod> tmpPaginObject = this.tQuestionMethodService
				.getPaginationObjectByParams4HQL(conditions, this.getPaginationParam());
		List<TQuestionMethod> question = new ArrayList<TQuestionMethod>();
		for (int i = 0; i < tmpPaginObject.getResultList().size(); i++) {
			TQuestionMethod qu = tmpPaginObject.getResultList().get(i);
			TQuestionMethod tQuestionMethod = (TQuestionMethod) this.tQuestionMethodService.getById(qu.getId());
			Map<String, Object> conditions2 = new HashMap<String, Object>();
			conditions2.put("id", qu.getId());
			PaginationObject<TImprovementMethod> tmpPaginObject2 = this.tQuestionMethodService
					.SearchTImprovementMethod(conditions2, this.getPaginationParam());
			tQuestionMethod.settImprovementMethod(tmpPaginObject2.getResultList());
			// 添加文字前的数字（加措施）
			for (int j = 0; j < tQuestionMethod.gettImprovementMethod().size(); j++) {
				TImprovementMethod method = tQuestionMethod.gettImprovementMethod().get(j);
				method.setMethod(j + 1 + "." + method.getMethod() + "（责任单位：" + method.getUnitName() + "；完成时间："
						+ method.getCompleteTime() + "）");
			}
			tQuestionMethod.setQuestion("（" + ToCH(i + 1) + "）" + tQuestionMethod.getQuestion());
			String[] strs = tQuestionMethod.getReason().split("\r\n");
			if (strs.length > 0) {
				String ss = "";
				for (int t = 0; t < strs.length; t++) {
					if (!strs[t].equals("")) {
						ss += strs[t]
								+ "</w:t></w:r></w:p><w:p wsp:rsidR=\"00C20E64\" wsp:rsidRPr=\"00A44351\" wsp:rsidRDefault=\"000D44CF\" wsp:rsidP=\"00C20E64\"><w:pPr><w:spacing w:line=\"560\" w:line-rule=\"exact\"/><w:ind w:first-line-chars=\"200\" w:first-line=\"640\"/><w:rPr><w:rFonts w:ascii=\"仿宋_GB2312\" w:fareast=\"仿宋_GB2312\"/><wx:font wx:val=\"仿宋_GB2312\"/><w:sz w:val=\"32\"/><w:sz-cs w:val=\"32\"/></w:rPr></w:pPr><w:r wsp:rsidRPr=\"00A44351\"><w:rPr><w:rFonts w:ascii=\"仿宋_GB2312\" w:fareast=\"仿宋_GB2312\" w:hint=\"fareast\"/><wx:font wx:val=\"仿宋_GB2312\"/><w:sz w:val=\"32\"/><w:sz-cs w:val=\"32\"/></w:rPr><w:t>";
					}
				}
				tQuestionMethod.setReason(ss);
			}
			question.add(tQuestionMethod);
		}
		return question;
	}

	// 获取上月整改数据
	private List<TImprovementUnit> getActData(String orgId, String catalog, String date) throws Exception {
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;
		Calendar cal = Calendar.getInstance();
		cal.setTime(sdf2.parse(date));
		cal.add(Calendar.MONTH, -1);
		// 上个月
		String lastMonth = sdf2.format(cal.getTime());
		List<TImprovementUnit> act = new ArrayList<TImprovementUnit>();
		Map<String, Object> conditions4 = new HashMap<String, Object>();
		conditions4.put("orgId", "");
		conditions4.put("catalog", catalog);
		conditions4.put("dataPeriod", lastMonth.replace("-", ""));
		conditions4.put("dutyOrgId", "");
		conditions4.put("questionOrgId", orgId);
		PaginationObject<TImprovementUnit> tmpPaginObject4 = this.tImprovementUnitService
				.searchtImprovementUnit(conditions4, this.getPaginationParam());
		List<TImprovementUnit> fiTrachargedetail = tmpPaginObject4.getResultList();
		// 中文数字并拼接新对象
		String tempstr = "";
		int x = 1;
		int y = 1;
		for (int i = 0; i < fiTrachargedetail.size(); i++) {
			TImprovementUnit tu = fiTrachargedetail.get(i);
			if (!tu.getQuestion().equals(tempstr)) {
				y = 1;
				tempstr = tu.getQuestion();
				tu.setQuestion("（" + ToCH(x) + "）关于“" + tu.getQuestion() + "”整改措施落实情况：");
				Map<String, Object> conditions5 = new HashMap<String, Object>();
				conditions5.put("improvementUnitid", tu.getId());
				PaginationObject<TImprovementActions> tmpPaginObject5 = this.tImprovementActionsService
						.searchttImprovementActions(conditions5, this.getPaginationParam());
				List<TImprovementActions> actions = tmpPaginObject5.getResultList();
				for (int j = 0; j < actions.size(); j++) {
					TImprovementActions action = actions.get(j);
					action.setActions(y + "." + action.getActions().split("。:")[0] + "。");
					y++;
				}
				tu.getProActions().addAll(actions);
				act.add(tu);
				x++;
			} else {
				for (TImprovementUnit unit : act) {
					if (unit.getQuestion().equals("（" + ToCH(x - 1) + "）关于“" + tu.getQuestion() + "”整改措施落实情况：")) {
						Map<String, Object> conditions5 = new HashMap<String, Object>();
						conditions5.put("improvementUnitid", tu.getId());
						PaginationObject<TImprovementActions> tmpPaginObject5 = this.tImprovementActionsService
								.searchttImprovementActions(conditions5, this.getPaginationParam());
						List<TImprovementActions> actions = tmpPaginObject5.getResultList();
						for (int j = 0; j < actions.size(); j++) {
							TImprovementActions action = actions.get(j);
							action.setActions(y + "." + action.getActions().split("（责任部门：")[0] + "。");
							y++;
						}
						unit.getProActions().addAll(actions);
					}
				}
			}
		}
		return act;
	}

	/**
	 * 财务部word导出 相关表格数据组装方法
	 *
	 */
	private void getData4CW(Map dataMap, String date, String orgId, String orgType, String qxtype, String deptname,
			String catalog, String dutyOrgId) throws Exception {
		int nowYear = Integer.parseInt(date.split("-")[0]);
		int nowMonth = Integer.parseInt(date.split("-")[1]);

		// 利润
		String lrindexId = "89156bc2-698e-415d-b288-0294b05dbe1f";
		String lrBYHBselId = "b51ba1fa-6a0b-4dd1-a39a-731a93a85d80";
		String lrQYDBselId = "3ed3fd19-aa92-40f7-929b-d45cb7a9501f";
		String lrJTDBselId = "20d4910a-339e-4e31-b039-8f86d403b618";
		String lrLSDBselId = "81711878-ee05-4c6f-a83e-ca6e477a80fe";
		// 单位容量利润
		String dwrllrindexId = "de755a9b-0347-4285-9458-a39b5ed3d3a1";
		String dwrllrBYHBselId = "1b924683-67a1-4f87-8c18-ce149de82383";
		String dwrllrQYDBselId = "1b924683-67a1-4f87-8c18-ce149de82384";
		String dwrllrJTDBselId = "1b924683-67a1-4f87-8c18-ce149de82385";
		String dwrllrLSDBselId = "1b924683-67a1-4f87-8c18-ce149de82386";
		// 度电利润
		String ddlrindexId = "d816b80e-2434-4b34-bab6-a95203866455";
		String ddlrBYHBselId = "1b924683-67a1-4f11-8c1b-ce149de82383";
		String ddlrQYDBselId = "1b924683-67a1-4f11-8c1b-ce149de82384";
		String ddlrJTDBselId = "1b924683-67a1-4f11-8c1b-ce149de82385";
		String ddlrLSDBselId = "1b924683-67a1-4f11-8c1b-ce149de82386";
		// 上网电价
		String swdjindexId = "e3206998-27bd-4ea4-bb4b-23cfa3de010c";
		String swdjBYHBselId = "57f7697f-1a07-4824-8c25-3eec43a4b9a7";
		String swdjQYDBselId = "2e4f7dc5-24c9-49dc-9180-39dba78434e4";
		String swdjJTDBselId = "65704ded-027b-4b34-965d-4b57a1879005";
		String swdjLSDBselId = "7dafcd97-b0f6-4b23-9828-e45a544c5a11";
		// 单位容量可控费用
		String dwrlkkfyindexId = "716dcccf-7021-4ba2-b003-8e8e7a02fa97";
		String dwrlkkfyBYHBselId = "1b924683-6733-4f87-8c1b-ce149de82383";
		String dwrlkkfyQYDBselId = "1b924683-6733-4f87-8c1b-ce149de82384";
		String dwrlkkfyJTDBselId = "1b924683-6733-4f87-8c1b-ce149de82385";
		String dwrlkkfyLSDBselId = "1b924683-6733-4f87-8c1b-ce149de82386";
		// 单位容量财务费用
		String dwrlcwfyindexId = "41d6d017-bcab-4708-9fab-06084ff7166d";
		String dwrlcwfyBYHBselId = "1b924683-61q1-4f87-8c1b-ce149de82383";
		String dwrlcwfyQYDBselId = "1b924683-61q1-4f87-8c1b-ce149de82384";
		String dwrlcwfyJTDBselId = "1b924683-61q1-4f87-8c1b-ce149de82385";
		String dwrlcwfyLSDBselId = "1b924683-61q1-4f87-8c1b-ce149de82386";

		Organization org = this.organMgrService.getOrganizationById(orgId);
		dataMap.put("orgName", org.getName());
		dataMap.put("year", String.valueOf(nowYear));
		dataMap.put("month", nowMonth);
		// 利润
		List<FZBenchmarkYTDVO> lrBYHBlists = getCWData4BYHB(orgId, lrindexId, date, lrBYHBselId);
		List<FZBenchmarkYTDVO> lrQYDBlists = getCWData4QYDB(orgId, lrindexId, date, lrQYDBselId);
		List<FZBenchmarkYTDVO> lrJTDBlists = getCWData4JTDB(orgId, lrindexId, date, lrJTDBselId);
		String lrMap = getCWData4LSDB(orgId, orgType, lrindexId, date, lrLSDBselId, "利润趋势图", "", "");
		// 利润
		dataMap.put("lrBYHBData", lrBYHBlists);
		dataMap.put("lrQYDBData", lrQYDBlists);
		dataMap.put("lrJTDBData", lrJTDBlists);
		dataMap.put("lrMap", lrMap);

		// 单位容量利润
		List<FZBenchmarkYTDVO> dwrllrBYHBlists = getCWData4BYHB(orgId, dwrllrindexId, date, dwrllrBYHBselId);
		List<FZBenchmarkYTDVO> dwrllrQYDBlists = getCWData4QYDB(orgId, dwrllrindexId, date, dwrllrQYDBselId);
		List<FZBenchmarkYTDVO> dwrllrJTDBlists = getCWData4JTDB(orgId, dwrllrindexId, date, dwrllrJTDBselId);
		String dwrllrMap = getCWData4LSDB(orgId, orgType, dwrllrindexId, date, dwrllrLSDBselId, "单位容量利润趋势图", "", "");
		// 单位容量利润
		dataMap.put("dwrllrBYHBData", dwrllrBYHBlists);
		dataMap.put("dwrllrQYDBData", dwrllrQYDBlists);
		dataMap.put("dwrllrJTDBData", dwrllrJTDBlists);
		dataMap.put("dwrllrMap", dwrllrMap);

		// 度电利润
		List<FZBenchmarkYTDVO> ddlrBYHBlists = getCWData4BYHB(orgId, ddlrindexId, date, ddlrBYHBselId);
		List<FZBenchmarkYTDVO> ddlrQYDBlists = getCWData4QYDB(orgId, ddlrindexId, date, ddlrQYDBselId);
		List<FZBenchmarkYTDVO> ddlrJTDBlists = getCWData4JTDB(orgId, ddlrindexId, date, ddlrJTDBselId);
		String ddlrMap = getCWData4LSDB(orgId, orgType, ddlrindexId, date, ddlrLSDBselId, "度电利润趋势图", "", "");
		// 度电利润
		dataMap.put("ddlrBYHBData", ddlrBYHBlists);
		dataMap.put("ddlrQYDBData", ddlrQYDBlists);
		dataMap.put("ddlrJTDBData", ddlrJTDBlists);
		dataMap.put("ddlrMap", ddlrMap);

		// 上网电价
		List<FZBenchmarkYTDVO> swdjBYHBlists = getCWData4BYHB(orgId, swdjindexId, date, swdjBYHBselId);
		List<FZBenchmarkYTDVO> swdjQYDBlists = getCWData4QYDB(orgId, swdjindexId, date, swdjQYDBselId);
		List<FZBenchmarkYTDVO> swdjJTDBlists = getCWData4JTDB(orgId, swdjindexId, date, swdjJTDBselId);
		String swdjMap = getCWData4LSDB(orgId, orgType, swdjindexId, date, swdjLSDBselId, "上网电价（含税）趋势图", "", "");
		// 上网电价
		dataMap.put("swdjBYHBData", swdjBYHBlists);
		dataMap.put("swdjQYDBData", swdjQYDBlists);
		dataMap.put("swdjJTDBData", swdjJTDBlists);
		dataMap.put("swdjMap", swdjMap);

		// 单位容量可控费用
		List<FZBenchmarkYTDVO> dwrlkkfyBYHBlists = getCWData4BYHB(orgId, dwrlkkfyindexId, date, dwrlkkfyBYHBselId);
		List<FZBenchmarkYTDVO> dwrlkkfyQYDBlists = getCWData4QYDB(orgId, dwrlkkfyindexId, date, dwrlkkfyQYDBselId);
		List<FZBenchmarkYTDVO> dwrlkkfyJTDBlists = getCWData4JTDB(orgId, dwrlkkfyindexId, date, dwrlkkfyJTDBselId);
		String dwrlkkfyMap = getCWData4LSDB(orgId, orgType, dwrlkkfyindexId, date, dwrlkkfyLSDBselId, "单位容量可控费用趋势图", "",
				"");
		// 单位容量可控费用
		dataMap.put("dwrlkkfyBYHBData", dwrlkkfyBYHBlists);
		dataMap.put("dwrlkkfyQYDBData", dwrlkkfyQYDBlists);
		dataMap.put("dwrlkkfyJTDBData", dwrlkkfyJTDBlists);
		dataMap.put("dwrlkkfyMap", dwrlkkfyMap);

		// 单位容量财务费用
		List<FZBenchmarkYTDVO> dwrlcwfyBYHBlists = getCWData4BYHB(orgId, dwrlcwfyindexId, date, dwrlcwfyBYHBselId);
		List<FZBenchmarkYTDVO> dwrlcwfyQYDBlists = getCWData4QYDB(orgId, dwrlcwfyindexId, date, dwrlcwfyQYDBselId);
		List<FZBenchmarkYTDVO> dwrlcwfyJTDBlists = getCWData4JTDB(orgId, dwrlcwfyindexId, date, dwrlcwfyJTDBselId);
		String dwrlcwfyMap = getCWData4LSDB(orgId, orgType, dwrlcwfyindexId, date, dwrlcwfyLSDBselId, "单位容量财务费用趋势图", "",
				"");
		// 单位容量财务费用
		dataMap.put("dwrlcwfyBYHBData", dwrlcwfyBYHBlists);
		dataMap.put("dwrlcwfyQYDBData", dwrlcwfyQYDBlists);
		dataMap.put("dwrlcwfyJTDBData", dwrlcwfyJTDBlists);
		dataMap.put("dwrlcwfyMap", dwrlcwfyMap);

		// 问题措施
		List<TQuestionMethod> question = getQuestionData(qxtype, deptname, catalog, date);
		dataMap.put("question", question);
		// 上月整改措施落实情况
		List<TImprovementUnit> act = getActData(orgId, catalog, date);
		dataMap.put("act", act);
	}

	// 获取本月环比数据
	private List<FZBenchmarkYTDVO> getCWData4BYHB(String orgId, String indexId, String date, String selId)
			throws Exception {
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;

		// ------------------表格数据---------------------
		List<FZBenchmarkYTDVO> tablelist = new ArrayList<FZBenchmarkYTDVO>();
		List<IndexProfit> tablelist1 = indexClassService.getMonthOneData(indexId, date, selId);
		// 环比
		Calendar ycal = Calendar.getInstance();
		ycal.setTime(sdf2.parse(date));
		ycal.add(Calendar.MONTH, -1);
		// 上个月
		String ynewDate = sdf2.format(ycal.getTime());
		List<IndexProfit> tablelist3 = indexClassService.getMonthOneData(indexId, ynewDate, selId);
		// 插入平均值
		java.text.DecimalFormat df = new java.text.DecimalFormat("#.##"); // 小数位数格式化两位
		if (null != tablelist1 && tablelist1.size() > 0) {
			for (int i = 0; i < tablelist1.size(); i++) {
				double indexVlaue = 0.0;
				double indexVlaue1 = 0.0;
				if (tablelist1.get(i).getAchieveValue() != null && tablelist1.get(i).getAchieveValue() != "") {
					indexVlaue = Double.valueOf(tablelist1.get(i).getAchieveValue());
				}
				if (null != tablelist3 && tablelist3.size() > 0) {
					for (int k = 0; k < tablelist3.size(); k++) {

						if (tablelist1.get(i).getSysId().equals(tablelist3.get(k).getSysId())) {
							if (tablelist3.get(k).getAchieveValue() != null
									&& tablelist3.get(k).getAchieveValue() != "") {
								indexVlaue1 = Double.valueOf(tablelist3.get(k).getAchieveValue());
							}
						}
					}
				}
				if (tablelist1.get(i).getSysId().equals(orgId)) {
					tablelist.add(new FZBenchmarkYTDVO(tablelist1.get(i).getOrgName(),
							String.valueOf(df.format(indexVlaue)), String.valueOf(df.format(indexVlaue1)),
							String.valueOf(df.format(indexVlaue - indexVlaue1)), "", "", "", "", "", "", 0, "", "",
							fillcolor));
				} else {
					tablelist.add(new FZBenchmarkYTDVO(tablelist1.get(i).getOrgName(),
							String.valueOf(df.format(indexVlaue)), String.valueOf(df.format(indexVlaue1)),
							String.valueOf(df.format(indexVlaue - indexVlaue1)), "", "", "", "", "", "", 0, "", "",
							fillcolorNO));
				}
			}
		}
		return tablelist;
	}

	// 获取区域对标数据
	private List<FZBenchmarkYTDVO> getCWData4QYDB(String orgId, String indexId, String date, String selId)
			throws Exception {
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;
		java.text.DecimalFormat df = new java.text.DecimalFormat("#.##"); // 小数位数格式化两位
		// ------------------表格数据---------------------
		List<FZBenchmarkYTDVO> tablelist = new ArrayList<FZBenchmarkYTDVO>();

		List<IndexProfit> tablelistavg = IndexCWService.getTableDataAVG(indexId, date, selId);

		List<IndexProfit> tablelist1 = IndexCWService.getTableData(indexId, date, selId);
		if (tablelistavg.size() > 0) {
			String orgName = "区域平均";
			String monthAvg = "";
			String yearAvg = "";
			if (tablelistavg.get(0).getOrgName() != null && tablelistavg.get(0).getOrgName() != "") {
				orgName = tablelistavg.get(0).getOrgName();
			}
			if (tablelistavg.get(0).getMonthAvgValue() != null && tablelistavg.get(0).getMonthAvgValue() != "") {
				monthAvg = tablelistavg.get(0).getMonthAvgValue();
			}
			if (tablelistavg.get(0).getYearAvgValue() != null && tablelistavg.get(0).getYearAvgValue() != "") {
				yearAvg = tablelistavg.get(0).getYearAvgValue();
			}
			tablelist.add(new FZBenchmarkYTDVO(orgName, monthAvg, "-", yearAvg, "-", "", "", "", "", "", -1, "", "",
					fillcolorNO));
		}
		// 修改背景色
		if (orgId.equals("40288f9b542d388801542d5099f90006") || orgId.equals("40288f9b542d388801542d50ea620007")
				|| orgId.equals("40288f9b542d388801542d51ea520009")
				|| orgId.equals("40288f9b542d388801542d51913d0008")) {
			for (int i = 0; i < tablelist1.size(); i++) {
				if (!tablelist1.get(i).getSysId().equals(orgId)) {
					tablelist1.get(i).setAnnualObject("其他");
				}
			}
		}

		for (int i = 0; i < tablelist1.size(); i++) {
			double monthVlaue = 0.0;
			double yearVlaue1 = 0.0;
			int sortNo = 0;
			if (tablelist1.get(i).getYearSortNO() != null && tablelist1.get(i).getYearSortNO() != "") {
				Integer it = new Integer(tablelist1.get(i).getYearSortNO());
				sortNo = it.intValue();
			}
			if (tablelist1.get(i).getAchieveValue() != null && tablelist1.get(i).getAchieveValue() != "") {
				monthVlaue = Double.valueOf(tablelist1.get(i).getAchieveValue());
			}
			if (tablelist1.get(i).getAddupValue() != null && tablelist1.get(i).getAddupValue() != "") {
				yearVlaue1 = Double.valueOf(tablelist1.get(i).getAddupValue());
			}
			if ("大唐".equals(tablelist1.get(i).getAnnualObject())) {
				tablelist
						.add(new FZBenchmarkYTDVO(tablelist1.get(i).getOrgName(), String.valueOf(df.format(monthVlaue)),
								tablelist1.get(i).getMonthSortNO(), String.valueOf(df.format(yearVlaue1)),
								tablelist1.get(i).getYearSortNO(), "", "", "", "", "", sortNo, "", "", fillcolor));
			} else {
				tablelist
						.add(new FZBenchmarkYTDVO(tablelist1.get(i).getOrgName(), String.valueOf(df.format(monthVlaue)),
								tablelist1.get(i).getMonthSortNO(), String.valueOf(df.format(yearVlaue1)),
								tablelist1.get(i).getYearSortNO(), "", "", "", "", "", sortNo, "", "", fillcolorNO));
			}
		}
		return tablelist;
	}

	// 获取集团对标数据
	private List<FZBenchmarkYTDVO> getCWData4JTDB(String orgId, String indexId, String date, String selId)
			throws Exception {
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;
		java.text.DecimalFormat df = new java.text.DecimalFormat("#.##"); // 小数位数格式化两位
		List<FZBenchmarkYTDVO> tablelist = new ArrayList<FZBenchmarkYTDVO>();
		// ------------------表格数据---------------------
		List<CwTmodeDTO> tablelist1 = IndexCWService.getTableDataOUTFZ(indexId, date, selId);
		for (int i = 0; i < tablelist1.size(); i++) {
			double monthVlaue = 0.0;
			double yearVlaue1 = 0.0;
			int sortNo = 0;
			String yueNo = "";
			if (tablelist1.get(i).getNianno() != null && tablelist1.get(i).getNianno() != "") {
				Integer it = new Integer(tablelist1.get(i).getNianno());
				sortNo = it.intValue();
			}
			if (tablelist1.get(i).getYueno() != null && tablelist1.get(i).getYueno() != "") {
				yueNo = tablelist1.get(i).getYueno();
			}
			if (tablelist1.get(i).getYuezhi() != null && tablelist1.get(i).getYuezhi() != "") {
				monthVlaue = Double.valueOf(tablelist1.get(i).getYuezhi());
			}
			if (tablelist1.get(i).getYuezhi() != null && tablelist1.get(i).getYuezhi() != "") {
				monthVlaue = Double.valueOf(tablelist1.get(i).getYuezhi());
			}
			if (tablelist1.get(i).getnianzhi() != null && tablelist1.get(i).getnianzhi() != "") {
				yearVlaue1 = Double.valueOf(tablelist1.get(i).getnianzhi());
			}
			if ("大唐".equals(tablelist1.get(i).getCpname())) {
				tablelist.add(new FZBenchmarkYTDVO(tablelist1.get(i).getCpname(), String.valueOf(df.format(monthVlaue)),
						yueNo, String.valueOf(df.format(yearVlaue1)), tablelist1.get(i).getNianno(), "", "", "", "", "",
						sortNo, "", "", fillcolor));
			} else {
				tablelist.add(new FZBenchmarkYTDVO(tablelist1.get(i).getCpname(), String.valueOf(df.format(monthVlaue)),
						yueNo, String.valueOf(df.format(yearVlaue1)), tablelist1.get(i).getNianno(), "", "", "", "", "",
						sortNo, "", "", fillcolorNO));
			}
		}
		// 排序
		MyComparator myCompatator = new MyComparator();
		Collections.sort(tablelist, myCompatator);
		return tablelist;
	}

	// 获取历史对标图
	private String getCWData4LSDB(String orgId, String orgType, String indexId, String date, String selId,
			String chartTitle, String xName, String yName) throws Exception {
		String rtnStr = "";
		SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy-MM");
		String fdate = "";
		if (StringUtils.isNotBlank(date)) {
			fdate = sdf2.format(sdf2.parse(date));
		}
		date = fdate;
		java.text.DecimalFormat df = new java.text.DecimalFormat("#.##"); // 小数位数格式化两位
		int nowYear = Integer.parseInt(date.split("-")[0]);
		int nowMonth = Integer.parseInt(date.split("-")[1]);
		String newMonth = date.split("-")[1];
		// String dateStr = String.valueOf(nowYear) + "年 1--" +
		// String.valueOf(nowMonth) + "月";
		String dateStr = String.valueOf(nowYear) + "年";

		List<FZBenchmarkYTDVO> tablelist = new ArrayList<FZBenchmarkYTDVO>();
		List<IndexProfit> tablelist1 = null;

		if (orgType.equals("1")) {
			tablelist1 = indexFZBenchmarkYTDService.getIndexProfitOverYearDataByFZXY(indexId, newMonth, orgId);
		} else {
			tablelist1 = indexFZBenchmarkYTDService.getIndexProfitOverYearDataByJCXY(indexId, newMonth, orgId);
		}

		String effectCycle = "";
		if (tablelist1.size() > 0) {
			// 循环将数据加到tablelist
			for (int i = 0; i < tablelist1.size(); i++) {
				double indexVlaue = 0.0;
				if (tablelist1.get(i).getAddupValue() != null && tablelist1.get(i).getAddupValue() != "") {
					indexVlaue = Double.valueOf(tablelist1.get(i).getAddupValue());
				}
				effectCycle = tablelist1.get(i).getEffectCycle();
				if (effectCycle != null && effectCycle != "") {
					int sortNo = Integer.parseInt(effectCycle.split("-")[0]);
					if (effectCycle.equals(date)) {
						tablelist.add(new FZBenchmarkYTDVO(dateStr, String.valueOf(df.format(indexVlaue)), "", "", "",
								"", sortNo, ""));
					} else {
						tablelist.add(new FZBenchmarkYTDVO(effectCycle.split("-")[0] + "年",
								String.valueOf(df.format(indexVlaue)), "", "", "", "", sortNo, ""));
					}
				}
			}
			// 排序
			MyComparator myCompatator = new MyComparator();
			Collections.sort(tablelist, myCompatator);

			DefaultCategoryDataset dataset = new DefaultCategoryDataset();
			for (int i = 0; i < tablelist.size(); i++) {
				FZBenchmarkYTDVO index = tablelist.get(i);
				dataset.addValue(Double.parseDouble(index.getIndexValue()), "", index.getOrgName());
			}
			rtnStr = createLineChart(dataset, "", "", chartTitle, indexId + ".png", false, null);

		}
		return rtnStr;
	}

	// -------------------------------------工具方法----------------------------------------------------
	public static String ToCH(int intInput) {
		String si = String.valueOf(intInput);
		String sd = "";
		if (si.length() == 1) // 個
		{
			sd += GetCH(intInput);
			return sd;
		} else if (si.length() == 2)// 十
		{
			if (si.substring(0, 1).equals("1"))
				sd += "十";
			else
				sd += (GetCH(intInput / 10) + "十");
			sd += ToCH(intInput % 10);
		} else if (si.length() == 3)// 百
		{
			sd += (GetCH(intInput / 100) + "百");
			if (String.valueOf(intInput % 100).length() < 2)
				sd += "零";
			sd += ToCH(intInput % 100);
		} else if (si.length() == 4)// 千
		{
			sd += (GetCH(intInput / 1000) + "千");
			if (String.valueOf(intInput % 1000).length() < 3)
				sd += "零";
			sd += ToCH(intInput % 1000);
		} else if (si.length() == 5)// 萬
		{
			sd += (GetCH(intInput / 10000) + "萬");
			if (String.valueOf(intInput % 10000).length() < 4)
				sd += "零";
			sd += ToCH(intInput % 10000);
		}

		return sd;
	}

	private static String GetCH(int input) {
		String sd = "";
		switch (input) {
		case 1:
			sd = "一";
			break;
		case 2:
			sd = "二";
			break;
		case 3:
			sd = "三";
			break;
		case 4:
			sd = "四";
			break;
		case 5:
			sd = "五";
			break;
		case 6:
			sd = "六";
			break;
		case 7:
			sd = "七";
			break;
		case 8:
			sd = "八";
			break;
		case 9:
			sd = "九";
			break;
		default:
			break;
		}
		return sd;
	}

	private static int chineseNumber2Int(String chineseNumber) {
		int result = 0;
		int temp = 1;// 存放一个单位的数字如：十万
		int count = 0;// 判断是否有chArr
		char[] cnArr = new char[] { '一', '二', '三', '四', '五', '六', '七', '八', '九' };
		char[] chArr = new char[] { '十', '百', '千', '万', '亿' };
		for (int i = 0; i < chineseNumber.length(); i++) {
			boolean b = true;// 判断是否是chArr
			char c = chineseNumber.charAt(i);
			for (int j = 0; j < cnArr.length; j++) {// 非单位，即数字
				if (c == cnArr[j]) {
					if (0 != count) {// 添加下一个单位之前，先把上一个单位值添加到结果中
						result += temp;
						temp = 1;
						count = 0;
					}
					// 下标+1，就是对应的值
					temp = j + 1;
					b = false;
					break;
				}
			}
			if (b) {// 单位{'十','百','千','万','亿'}
				for (int j = 0; j < chArr.length; j++) {
					if (c == chArr[j]) {
						switch (j) {
						case 0:
							temp *= 10;
							break;
						case 1:
							temp *= 100;
							break;
						case 2:
							temp *= 1000;
							break;
						case 3:
							temp *= 10000;
							break;
						case 4:
							temp *= 100000000;
							break;
						default:
							break;
						}
						count++;
					}
				}
			}
			if (i == chineseNumber.length() - 1) {// 遍历到最后一个字符
				result += temp;
			}
		}
		return result;
	}

	// 补0操作
	public static String fillZero(int i) {
		String str = "";
		if (i > 0 && i < 10) {
			str = "0" + i;
		} else {
			str = "" + i;
		}
		return str;
	}

	// 将数字转化为大写
	public static String numToUpper(int num) {
		String u[] = { "〇", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
		char[] str = String.valueOf(num).toCharArray();
		String rstr = "";
		for (int i = 0; i < str.length; i++) {
			rstr = rstr + u[Integer.parseInt(str[i] + "")];
		}
		return rstr;
	}

	// 月转化为大写
	public static String monthToUppder(int month) {
		if (month < 10) {
			return numToUpper(month) + "月";
		} else if (month == 10) {
			return "十" + "月";
		} else {
			return "十" + numToUpper(month - 10) + "月";
		}
	}

	/**
	 * 柱状图
	 * 
	 * @param dataset
	 *            数据集
	 * @param xName
	 *            x轴的说明（如种类，时间等）
	 * @param yName
	 *            y轴的说明（如速度，时间等）
	 * @param chartTitle
	 *            图标题
	 * @param charName
	 *            生成图片的名字
	 * @return
	 */
	public static String createBarChart(DefaultCategoryDataset dataset, String xName, String yName, String chartTitle,
			String charName, boolean isPaint, String[] colors) {
		JFreeChart chart = ChartFactory.createBarChart(chartTitle, // 图表标题
				xName, // 目录轴的显示标签
				yName, // 数值轴的显示标签
				dataset, // 数据集
				PlotOrientation.VERTICAL, // 图表方向：水平、垂直
				true, // 是否显示图例(对于简单的柱状图必须是false)
				false, // 是否生成工具
				false // 是否生成URL链接
		);
		ChartUtils.setAntiAlias(chart);

		StandardChartTheme theme = new StandardChartTheme("unicode") {
			// 重写apply(...)方法是为了借机消除文字锯齿.VALUE_TEXT_ANTIALIAS_OFF
			public void apply(JFreeChart chart) {
				chart.getRenderingHints().put(RenderingHints.KEY_TEXT_ANTIALIASING,
						RenderingHints.VALUE_TEXT_ANTIALIAS_OFF);
				super.apply(chart);
			}
		};
		theme.setExtraLargeFont(new Font("宋体", Font.PLAIN, 20));
		theme.setLargeFont(new Font("宋体", Font.PLAIN, 14));
		theme.setRegularFont(new Font("宋体", Font.PLAIN, 12));
		theme.setSmallFont(new Font("宋体", Font.PLAIN, 10));

		ChartFactory.setChartTheme(theme);

		Font labelFont = new Font("SansSerif", Font.TRUETYPE_FONT, 16);
		/*
		 * VALUE_TEXT_ANTIALIAS_OFF表示将文字的抗锯齿关闭,
		 * 使用的关闭抗锯齿后，字体尽量选择12到14号的宋体字,这样文字最清晰好看
		 */
		// chart.getRenderingHints().put(RenderingHints.KEY_TEXT_ANTIALIASING,RenderingHints.VALUE_TEXT_ANTIALIAS_OFF);
		chart.setTextAntiAlias(false);
		chart.setBorderVisible(false);
		chart.setBackgroundPaint(Color.white);
		// 处理主标题的乱码
		chart.getTitle().setFont(new Font("宋体", Font.BOLD, 20));
		// 处理子标题乱码
		// chart.getLegend().setItemFont(new Font("宋体",Font.BOLD,15));
		chart.getLegend().setVisible(isPaint);// 显示或隐藏底部文字
		// create plot
		CategoryPlot plot = chart.getCategoryPlot();
		// 设置横虚线可见
		plot.setRangeGridlinesVisible(true);
		// 虚线色彩
		plot.setRangeGridlinePaint(Color.gray);
		// 数据轴精度
		NumberAxis vn = (NumberAxis) plot.getRangeAxis();
		// vn.setAutoRangeIncludesZero(true);
		// DecimalFormat df = new DecimalFormat("#0.00");
		// vn.setNumberFormatOverride(df); // 数据轴数据标签的显示格式
		// x轴设置

		CategoryAxis domainAxis = plot.getDomainAxis();
		domainAxis.setLabelFont(labelFont);// 轴标题
		domainAxis.setLabelPaint(Color.BLACK);
		domainAxis.setTickMarkPaint(Color.BLACK);
		domainAxis.setTickLabelFont(labelFont);// 轴数值
		// Lable（Math.PI/3.0）度倾斜
		// domainAxis.setCategoryLabelPositions(CategoryLabelPositions
		// .createUpRotationLabelPositions(Math.PI / 3.0));
		domainAxis.setMaximumCategoryLabelWidthRatio(0.6f);// 横轴上的 Lable 是否完整显示

		// 设置距离图片左端距离
		domainAxis.setLowerMargin(0.02);
		// 设置距离图片右端距离
		domainAxis.setUpperMargin(0.02);
		// 设置 columnKey 是否间隔显示
		// domainAxis.setSkipCategoryLabelsToFit(true);
		plot.setDomainAxis(domainAxis);
		// 设置柱图背景色（注意，系统取色的时候要使用16位的模式来查看颜色编码，这样比较准确）

		CategoryAxis categoryAxis = plot.getDomainAxis();
		// 横轴上的 Lable 90度倾斜
		categoryAxis.setCategoryLabelPositions(CategoryLabelPositions.UP_45);

		plot.setBackgroundPaint(Color.WHITE);
		// y轴设置
		ValueAxis rangeAxis = plot.getRangeAxis();
		rangeAxis.setLabelFont(labelFont);
		rangeAxis.setTickLabelFont(labelFont);
		// 设置最高的一个 Item 与图片顶端的距离
		rangeAxis.setUpperMargin(0.15);
		// 设置最低的一个 Item 与图片底端的距离
		rangeAxis.setLowerMargin(0.15);
		plot.setRangeAxis(rangeAxis);
		BarRenderer renderer;
		if (null != colors && colors.length > 0) {// 如果传进去颜色则显示否则按三个默认颜色分类
			renderer = new CustomRenderer(colors);
		} else {
			renderer = new BarRenderer();
			renderer.setSeriesPaint(0, Color.decode("#4F81BD")); // 给series2 Bar
			renderer.setSeriesPaint(1, Color.decode("#C0504D")); // 给series1 Bar
			renderer.setSeriesPaint(2, Color.decode("#9BBB59")); // 给series3 Bar
		}
		// 设置柱子宽度

		if (dataset.getColumnCount() > 10) {
			renderer.setMaximumBarWidth(0.025);
		} else {
			renderer.setMaximumBarWidth(0.045);
		}

		// 设置柱子高度
		renderer.setMinimumBarLength(0.2);
		// 设置柱子边框颜色
		// renderer.setBaseOutlinePaint(Color.BLACK);
		// 设置柱子边框可见
		renderer.setDrawBarOutline(false);

		// 显示每个柱的数值，并修改该数值的字体属性
		renderer.setIncludeBaseInRange(true);
		renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
		renderer.setBaseItemLabelsVisible(true);
		renderer.setBarPainter(new StandardBarPainter());
		// 设置每个地区所包含的平行柱的之间距离
		renderer.setItemMargin(-0.01);
		renderer.setShadowVisible(false);
		renderer.setBaseItemLabelFont(new Font("SansSerif", Font.TRUETYPE_FONT, 14));

		plot.setRenderer(renderer);
		// 设置柱的透明度
		plot.setForegroundAlpha(1.0f);
		BASE64Encoder BASE64 = new BASE64Encoder();
		ByteArrayOutputStream bas = new ByteArrayOutputStream();
		try {
			ChartUtilities.writeChartAsJPEG(bas, 1.0f, chart, 850, 300, null);
			bas.flush();
			bas.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		byte[] byteArray = bas.toByteArray();
		try {
			InputStream is = new ByteArrayInputStream(byteArray);
			byteArray = new byte[is.available()];
			is.read(byteArray);
			is.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		String xml_img = BASE64.encode(byteArray);
		return xml_img;
	}

	/**
	 * 折线图
	 * 
	 * @param dataset
	 *            数据集
	 * @param xName
	 *            x轴的说明（如种类，时间等）
	 * @param yName
	 *            y轴的说明（如速度，时间等）
	 * @param chartTitle
	 *            图标题
	 * @param charName
	 *            生成图片的名字
	 * @return
	 */
	public static String createLineChart(DefaultCategoryDataset dataset, String xName, String yName, String chartTitle,
			String charName, boolean isPaint, String[] colors) {
		JFreeChart chart = ChartFactory.createLineChart(chartTitle, // 图表标题
				xName, // 目录轴的显示标签
				yName, // 数值轴的显示标签
				dataset, // 数据集
				PlotOrientation.VERTICAL, // 图表方向：水平、垂直
				true, // 是否显示图例(对于简单的柱状图必须是false)
				false, // 是否生成工具
				false // 是否生成URL链接
		);
		ChartUtils.setAntiAlias(chart);

		StandardChartTheme theme = new StandardChartTheme("unicode") {
			// 重写apply(...)方法是为了借机消除文字锯齿.VALUE_TEXT_ANTIALIAS_OFF
			public void apply(JFreeChart chart) {
				chart.getRenderingHints().put(RenderingHints.KEY_TEXT_ANTIALIASING,
						RenderingHints.VALUE_TEXT_ANTIALIAS_OFF);
				super.apply(chart);
			}
		};
		theme.setExtraLargeFont(new Font("宋体", Font.PLAIN, 20));
		theme.setLargeFont(new Font("宋体", Font.PLAIN, 14));
		theme.setRegularFont(new Font("宋体", Font.PLAIN, 12));
		theme.setSmallFont(new Font("宋体", Font.PLAIN, 10));

		ChartFactory.setChartTheme(theme);

		Font labelFont = new Font("SansSerif", Font.TRUETYPE_FONT, 16);
		/*
		 * VALUE_TEXT_ANTIALIAS_OFF表示将文字的抗锯齿关闭,
		 * 使用的关闭抗锯齿后，字体尽量选择12到14号的宋体字,这样文字最清晰好看
		 */
		chart.getRenderingHints().put(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_OFF);
		chart.setTextAntiAlias(false);
		chart.setBorderVisible(false);
		chart.setBackgroundPaint(Color.white);
		// 处理主标题的乱码
		chart.getTitle().setFont(new Font("宋体", Font.BOLD, 20));
		// 处理子标题乱码
		// chart.getLegend().setItemFont(new Font("宋体",Font.BOLD,15));
		chart.getLegend().setVisible(isPaint);// 显示或隐藏底部文字
		// create plot
		CategoryPlot plot = chart.getCategoryPlot();
		// 设置横虚线可见
		plot.setRangeGridlinesVisible(true);
		// 虚线色彩
		plot.setRangeGridlinePaint(Color.gray);
		// 数据轴精度
		NumberAxis vn = (NumberAxis) plot.getRangeAxis();
		// vn.setAutoRangeIncludesZero(true);
		// DecimalFormat df = new DecimalFormat("#0.00");
		// vn.setNumberFormatOverride(df); // 数据轴数据标签的显示格式
		// x轴设置

		CategoryAxis domainAxis = plot.getDomainAxis();
		domainAxis.setLabelFont(labelFont);// 轴标题
		domainAxis.setLabelPaint(Color.BLACK);
		domainAxis.setTickMarkPaint(Color.BLACK);
		domainAxis.setTickLabelFont(labelFont);// 轴数值
		// Lable（Math.PI/3.0）度倾斜
		// domainAxis.setCategoryLabelPositions(CategoryLabelPositions
		// .createUpRotationLabelPositions(Math.PI / 3.0));
		domainAxis.setMaximumCategoryLabelWidthRatio(0.6f);// 横轴上的 Lable 是否完整显示

		// 设置距离图片左端距离
		domainAxis.setLowerMargin(0.02);
		// 设置距离图片右端距离
		domainAxis.setUpperMargin(0.02);
		// 设置 columnKey 是否间隔显示
		// domainAxis.setSkipCategoryLabelsToFit(true);
		plot.setDomainAxis(domainAxis);
		// 设置柱图背景色（注意，系统取色的时候要使用16位的模式来查看颜色编码，这样比较准确）

		CategoryAxis categoryAxis = plot.getDomainAxis();
		// 横轴上的 Lable 90度倾斜
		categoryAxis.setCategoryLabelPositions(CategoryLabelPositions.UP_45);

		plot.setBackgroundPaint(Color.WHITE);
		// y轴设置
		ValueAxis rangeAxis = plot.getRangeAxis();
		rangeAxis.setLabelFont(labelFont);
		rangeAxis.setTickLabelFont(labelFont);
		// 设置最高的一个 Item 与图片顶端的距离
		rangeAxis.setUpperMargin(0.15);
		// 设置最低的一个 Item 与图片底端的距离
		rangeAxis.setLowerMargin(0.15);
		plot.setRangeAxis(rangeAxis);
		LineAndShapeRenderer renderer;
		if (null != colors && colors.length > 0) {// 如果传进去颜色则显示否则按三个默认颜色分类
			renderer = new LineAndShapeRenderer();
		} else {
			renderer = new LineAndShapeRenderer();
			renderer.setSeriesPaint(0, Color.decode("#C0504D")); // 给series1 Bar
			renderer.setSeriesPaint(1, Color.decode("#4F81BD")); // 给series2 Bar
			renderer.setSeriesPaint(2, Color.decode("#9BBB59")); // 给series3 Bar
		}

		// 显示每个柱的数值，并修改该数值的字体属性
		// renderer.setIncludeBaseInRange(true);
		renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
		renderer.setBaseItemLabelsVisible(true);
		renderer.setBaseItemLabelFont(new Font("SansSerif", Font.TRUETYPE_FONT, 14));
		renderer.setSeriesStroke(0, new BasicStroke(4.0F));
		renderer.setSeriesStroke(1, new BasicStroke(4.0F));
		renderer.setStroke(new BasicStroke(1.5F));
		renderer.setBaseItemLabelsVisible(true);
		renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator(
				StandardCategoryItemLabelGenerator.DEFAULT_LABEL_FORMAT_STRING, NumberFormat.getInstance()));
		renderer.setBasePositiveItemLabelPosition(
				new ItemLabelPosition(ItemLabelAnchor.OUTSIDE1, TextAnchor.BOTTOM_CENTER));// weizhi
		renderer.setBaseShapesVisible(true);// 数据点绘制形状
		plot.setRenderer(renderer);
		// 设置柱的透明度
		plot.setForegroundAlpha(1.0f);
		BASE64Encoder BASE64 = new BASE64Encoder();
		ByteArrayOutputStream bas = new ByteArrayOutputStream();
		try {
			ChartUtilities.writeChartAsJPEG(bas, 1.0f, chart, 850, 300, null);
			bas.flush();
			bas.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		byte[] byteArray = bas.toByteArray();
		try {
			InputStream is = new ByteArrayInputStream(byteArray);
			byteArray = new byte[is.available()];
			is.read(byteArray);
			is.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		String xml_img = BASE64.encode(byteArray);
		return xml_img;
	}

	// 合并word文件
	public static File uniteDoc(List fileList, String savepaths) {
		if (fileList.size() == 0 || fileList == null) {
			return null;
		}
		// 打开word
		ActiveXComponent app = new ActiveXComponent("Word.Application");// 启动word
		try {
			// 设置word不可见
			app.setProperty("Visible", new Variant(false));
			// 获得documents对象
			Object docs = app.getProperty("Documents").toDispatch();
			// 打开第一个文件
			Object doc = Dispatch.invoke((Dispatch) docs, "Open", Dispatch.Method,
					new Object[] { (String) fileList.get(fileList.size() - 1), new Variant(false), new Variant(true) },
					new int[3]).toDispatch();

			File file = new File(savepaths);
			if (!file.exists()) {
				file.createNewFile();
			}
			// 追加文件
			for (int i = 0; i < fileList.size() - 1; i++) {
				Dispatch.invoke(app.getProperty("Selection").toDispatch(), "insertFile", Dispatch.Method, new Object[] {
						(String) fileList.get(i), "", new Variant(false), new Variant(false), new Variant(false) },
						new int[3]);
			}
			// 保存新的word文件
			Dispatch.invoke((Dispatch) doc, "SaveAs", Dispatch.Method, new Object[] { savepaths, new Variant(1) },
					new int[3]);
			Variant f = new Variant(false);
			Dispatch.call((Dispatch) doc, "Close", f);
			return file;
		} catch (Exception e) {
			throw new RuntimeException("合并word文件出错.原因:" + e);
		} finally {
			app.invoke("Quit", new Variant[] {});
		}
	}
}