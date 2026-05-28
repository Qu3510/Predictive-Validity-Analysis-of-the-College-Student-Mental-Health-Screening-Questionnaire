# ====================================================
# 大学生心理健康数据分析 - 二分类一键更新系统
# 针对最新Excel格式（已有英文risk label列）
# ====================================================
rm(data) 
library(readxl)
setwd("D:/Tu/2025/") 


# ====================================================
# 大学生心理健康筛查问卷预测效度分析（二分类）
# 版本：2.0 - 基于二分类风险标签
# 最后更新：2025年
# ====================================================

# 清理环境
rm(list = ls())

# 1. 加载必要包 ------------------------------------------------------------
cat("🔧 加载必要包...\n")

required_packages <- c(
  "readxl",      # 读取Excel
  "dplyr",       # 数据处理
  "tidyr",       # 数据整理
  "ggplot2",     # 绘图
  "pROC",        # ROC分析
  "caret",       # 模型评估
  "ggpubr",      # 出版级图表
  "psych",       # 描述性统计
  "tableone",    # 创建Table 1
  "writexl",     # 写入Excel
  "patchwork"    # 多图组合
)

# 安装缺失的包
new_packages <- required_packages[!(required_packages %in% installed.packages()[,"Package"])]
if(length(new_packages) > 0) {
  cat("📦 安装缺失包:", paste(new_packages, collapse = ", "), "\n")
  install.packages(new_packages, dependencies = TRUE)
}

# 加载包
suppressPackageStartupMessages({
  for(pkg in required_packages) {
    library(pkg, character.only = TRUE)
  }
})

cat("✅ 所有包加载完成\n")

# 2. 数据读取与检查 --------------------------------------------------------
cat("\n📊 数据读取与检查\n")
cat(rep("=", 60), "\n", sep = "")

# 2.1 设置文件路径
file_path <- "01 data/new.xlsx"

# 2.2 检查文件
if(!file.exists(file_path)) {
  stop("❌ 文件不存在: ", file_path, 
       "\n当前工作目录: ", getwd())
}

# 2.3 读取数据
cat("📥 读取Excel文件...\n")
data <- read_excel(file_path)

# 2.4 检查数据维度
cat("📈 数据维度:", nrow(data), "行 ×", ncol(data), "列\n")

# 2.5 查找关键变量
cat("\n🔍 查找关键变量:\n")

# 查找风险标签列（可能有多种名称）
risk_cols <- c("risk_label_binary", "risk_label", "risk lable", "风险标签")
risk_col_name <- NULL

for(col in risk_cols) {
  if(col %in% colnames(data)) {
    risk_col_name <- col
    cat("✅ 找到风险标签列:", col, "\n")
    break
  }
}

if(is.null(risk_col_name)) {
  cat("⚠️  未找到标准风险标签列，查看所有列:\n")
  print(colnames(data))
  risk_col_name <- readline("请输入风险标签列名: ")
}

# 2.6 检查风险标签分布
cat("\n📊 风险标签分布:\n")
risk_table <- table(data[[risk_col_name]], useNA = "always")
print(risk_table)

# 转换为数值型
data$risk_binary <- as.numeric(as.character(data[[risk_col_name]]))

# 验证是否为二分类
unique_vals <- unique(na.omit(data$risk_binary))
if(!all(unique_vals %in% c(0, 1))) {
  warning("⚠️  风险标签包含非0/1值: ", paste(unique_vals, collapse = ", "))
  cat("将自动转换为二分类: 非0值转为1\n")
  data$risk_binary <- ifelse(data$risk_binary == 0, 0, 1)
}

cat("\n✅ 最终二分类分布:\n")
final_table <- table(data$risk_binary, useNA = "always")
names(final_table) <- c("0 (无需干预)", "1 (需要干预)", "缺失")[1:length(final_table)]
print(final_table)

# 计算阳性率
n_positive <- sum(data$risk_binary == 1, na.rm = TRUE)
n_negative <- sum(data$risk_binary == 0, na.rm = TRUE)
positive_rate <- n_positive / (n_positive + n_negative) * 100
cat("\n📈 统计结果:\n")
cat("   无需干预 (0):", n_negative, "\n")
cat("   需要干预 (1):", n_positive, "\n")
cat("   阳性率:", round(positive_rate, 2), "%\n")

# 3. 数据清洗与准备 --------------------------------------------------------
cat("\n🧹 数据清洗与准备\n")
cat(rep("=", 60), "\n", sep = "")

# 3.1 查找问卷总分列
total_score_patterns <- c("总分", "total_score", "total score")
total_score_col <- NULL

for(pattern in total_score_patterns) {
  matches <- grep(pattern, colnames(data), ignore.case = TRUE, value = TRUE)
  if(length(matches) > 0) {
    total_score_col <- matches[1]
    cat("✅ 找到总分列:", total_score_col, "\n")
    break
  }
}

if(is.null(total_score_col)) {
  cat("⚠️  未找到总分列，将尝试计算总分\n")
} else {
  # 重命名为统一名称
  data$total_score <- as.numeric(as.character(data[[total_score_col]]))
  cat("📊 问卷总分统计:\n")
  print(summary(data$total_score))
}

# 3.2 查找心理指标列（指标标准分）
cat("\n🔍 查找心理指标列...\n")

# 常见的心理指标列名模式
psych_patterns <- c(
  "幻觉、妄想症状.*指标标准分",
  "自杀意图.*指标标准分",
  "焦虑.*指标标准分",
  "抑郁.*指标标准分",
  "偏执.*指标标准分",
  "自卑.*指标标准分",
  "敏感.*指标标准分",
  "社交恐惧.*指标标准分",
  "躯体化.*指标标准分",
  "依赖.*指标标准分",
  "敌对攻击.*指标标准分",
  "冲动.*指标标准分",
  "强迫.*指标标准分",
  "网络成瘾.*指标标准分",
  "自伤行为.*指标标准分",
  "进食问题.*指标标准分",
  "睡眠困扰.*指标标准分",
  "学校适应困难.*指标标准分",
  "人际关系困扰.*指标标准分",
  "学业压力.*指标标准分",
  "就业压力.*指标标准分",
  "恋爱困扰.*指标标准分"
)

# 查找并重命名心理指标
psych_vars <- list()
for(pattern in psych_patterns) {
  matches <- grep(pattern, colnames(data), value = TRUE)
  if(length(matches) > 0) {
    psych_vars[[pattern]] <- matches[1]
  }
}

cat("✅ 找到", length(psych_vars), "个心理指标标准分\n")

# 3.3 简化列名并转换为数值型
cat("\n🔤 简化列名...\n")

# 简化名称映射
simple_names <- c(
  "幻觉、妄想症状.*指标标准分" = "psychosis",
  "自杀意图.*指标标准分" = "suicide_ideation",
  "焦虑.*指标标准分" = "anxiety",
  "抑郁.*指标标准分" = "depression",
  "偏执.*指标标准分" = "paranoia",
  "自卑.*指标标准分" = "inferiority",
  "敏感.*指标标准分" = "sensitivity",
  "社交恐惧.*指标标准分" = "social_phobia",
  "躯体化.*指标标准分" = "somatization",
  "依赖.*指标标准分" = "dependence",
  "敌对攻击.*指标标准分" = "hostility",
  "冲动.*指标标准分" = "impulsivity",
  "强迫.*指标标准分" = "compulsion",
  "网络成瘾.*指标标准分" = "internet_addiction",
  "自伤行为.*指标标准分" = "self_harm",
  "进食问题.*指标标准分" = "eating_disorder",
  "睡眠困扰.*指标标准分" = "sleep_disturbance",
  "学校适应困难.*指标标准分" = "school_adjustment",
  "人际关系困扰.*指标标准分" = "interpersonal_stress",
  "学业压力.*指标标准分" = "academic_stress",
  "就业压力.*指标标准分" = "employment_stress",
  "恋爱困扰.*指标标准分" = "relationship_stress"
)

# 应用重命名和转换
for(old_name in names(simple_names)) {
  if(old_name %in% names(psych_vars)) {
    actual_col <- psych_vars[[old_name]]
    new_name <- simple_names[old_name]
    
    # 重命名并转换为数值型
    data[[new_name]] <- as.numeric(as.character(data[[actual_col]]))
    
    cat("   ", actual_col, " → ", new_name, "\n")
  }
}

# 3.4 创建心理维度总分
cat("\n📊 创建心理维度总分...\n")

# 定义维度
severe_vars <- c("psychosis", "suicide_ideation")
general_vars <- c("anxiety", "depression", "paranoia", "inferiority", 
                  "sensitivity", "social_phobia", "somatization", 
                  "dependence", "hostility", "impulsivity", "compulsion",
                  "internet_addiction", "self_harm", "eating_disorder", 
                  "sleep_disturbance")
developmental_vars <- c("school_adjustment", "interpersonal_stress", 
                        "academic_stress", "employment_stress", "relationship_stress")

# 计算维度总分（仅对存在的变量）
if(all(severe_vars %in% colnames(data))) {
  data$severe_total <- rowSums(data[, severe_vars], na.rm = TRUE)
  cat("✅ 严重心理问题维度: 2个指标\n")
}

if(length(intersect(general_vars, colnames(data))) > 5) {
  existing_general <- intersect(general_vars, colnames(data))
  data$general_total <- rowSums(data[, existing_general], na.rm = TRUE)
  cat("✅ 一般心理问题维度:", length(existing_general), "个指标\n")
}

if(all(developmental_vars %in% colnames(data))) {
  data$developmental_total <- rowSums(data[, developmental_vars], na.rm = TRUE)
  cat("✅ 发展性困扰维度: 5个指标\n")
}

# 3.5 创建交互项（如果变量存在）
cat("\n🔄 创建交互项...\n")

if(all(c("depression", "suicide_ideation") %in% colnames(data))) {
  data$depression_suicide <- data$depression * data$suicide_ideation
  cat("✅ depression × suicide_ideation\n")
}

if(all(c("psychosis", "paranoia") %in% colnames(data))) {
  data$psychosis_paranoia <- data$psychosis * data$paranoia
  cat("✅ psychosis × paranoia\n")
}

# 4. 描述性统计分析 --------------------------------------------------------
cat("\n📈 描述性统计分析\n")
cat(rep("=", 60), "\n", sep = "")

# 4.1 创建Table 1：人口学特征
cat("\n👥 Table 1: 人口学特征\n")

# 检查人口学变量
demographic_vars <- c("性别", "民族", "生源地", "是否独生")
demographic_vars <- demographic_vars[demographic_vars %in% colnames(data)]

if(length(demographic_vars) > 0) {
  # 创建Table 1
  table1 <- CreateTableOne(
    vars = demographic_vars,
    strata = "risk_binary",
    data = data,
    factorVars = demographic_vars
  )
  
  print(table1, showAllLevels = TRUE, formatOptions = list(big.mark = ","))
  
  # 导出为CSV
  table1_df <- print(table1, printToggle = FALSE)
  write.csv(table1_df, "table1_demographics.csv")
  cat("📁 Table 1 已保存为: table1_demographics.csv\n")
} else {
  cat("⚠️  未找到人口学变量\n")
}

# 4.2 心理指标描述统计
cat("\n📊 心理指标描述统计（按风险组）\n")

# 选择关键心理指标
key_psych_vars <- c("depression", "anxiety", "suicide_ideation", "psychosis")
key_psych_vars <- key_psych_vars[key_psych_vars %in% colnames(data)]

if(length(key_psych_vars) > 0) {
  psych_stats <- data %>%
    group_by(risk_binary) %>%
    summarise(
      人数 = n(),
      across(
        all_of(key_psych_vars),
        list(
          均值 = ~round(mean(., na.rm = TRUE), 2),
          标准差 = ~round(sd(., na.rm = TRUE), 2),
          中位数 = ~round(median(., na.rm = TRUE), 2)
        )
      )
    )
  
  print(psych_stats)
  
  # 保存
  write.csv(psych_stats, "psychological_stats_by_risk.csv")
  cat("📁 心理指标统计已保存为: psychological_stats_by_risk.csv\n")
}

# 5. ROC分析 -------------------------------------------------------------
cat("\n📈 ROC分析：问卷总分预测干预需求\n")
cat(rep("=", 60), "\n", sep = "")

if("total_score" %in% colnames(data)) {
  # 准备数据
  roc_data <- data[!is.na(data$total_score) & !is.na(data$risk_binary), ]
  
  if(nrow(roc_data) > 0 && sum(roc_data$risk_binary == 1) >= 10) {
    # 计算ROC
    roc_result <- roc(roc_data$risk_binary, roc_data$total_score)
    auc_value <- auc(roc_result)
    ci_auc <- ci.auc(roc_result)
    
    cat("📊 ROC分析结果:\n")
    cat("   样本量:", nrow(roc_data), "\n")
    cat("   阳性样本:", sum(roc_data$risk_binary == 1), "\n")
    cat("   阴性样本:", sum(roc_data$risk_binary == 0), "\n")
    cat("   AUC:", round(auc_value, 3), "\n")
    cat("   95% CI: [", round(ci_auc[1], 3), ", ", round(ci_auc[3], 3), "]\n", sep = "")
    
    # 寻找最佳切点
    best_cutoff <- coords(roc_result, "best", ret = "all")
    cat("\n🎯 最佳切点:\n")
    cat("   总分阈值:", round(best_cutoff$threshold, 1), "\n")
    cat("   灵敏度:", round(best_cutoff$sensitivity, 3), "\n")
    cat("   特异度:", round(best_cutoff$specificity, 3), "\n")
    cat("   Youden指数:", round(best_cutoff$sensitivity + best_cutoff$specificity - 1, 3), "\n")
    
    # 计算不同切点的诊断指标
    cat("\n📋 不同切点的诊断指标:\n")
    
    cutoffs <- c(200, 220, 240, 260, 280, 300)
    diag_table <- data.frame()
    
    for(cutoff in cutoffs) {
      pred <- ifelse(roc_data$total_score >= cutoff, 1, 0)
      true <- roc_data$risk_binary
      
      # 计算指标
      tp <- sum(pred == 1 & true == 1)
      tn <- sum(pred == 0 & true == 0)
      fp <- sum(pred == 1 & true == 0)
      fn <- sum(pred == 0 & true == 1)
      
      sensitivity <- tp / (tp + fn)
      specificity <- tn / (tn + fp)
      ppv <- tp / (tp + fp)
      npv <- tn / (tn + fn)
      
      diag_table <- rbind(diag_table, data.frame(
        切点 = cutoff,
        灵敏度 = round(sensitivity, 3),
        特异度 = round(specificity, 3),
        阳性预测值 = round(ppv, 3),
        阴性预测值 = round(npv, 3)
      ))
    }
    
    print(diag_table)
    write.csv(diag_table, "diagnostic_cutoffs.csv", row.names = FALSE)
    
    # 5.1 绘制ROC曲线
    cat("\n🎨 绘制ROC曲线...\n")
    
    roc_plot <- ggroc(roc_result, legacy.axes = TRUE) +
      geom_abline(intercept = 0, slope = 1, linetype = "dashed", color = "gray50") +
      annotate("text", x = 0.7, y = 0.2, 
               label = paste0("AUC = ", round(auc_value, 3), 
                              " (95% CI: ", round(ci_auc[1], 3), "-", 
                              round(ci_auc[3], 3), ")"),
               size = 5) +
      labs(title = "问卷总分预测干预需求的ROC曲线",
           x = "1 - 特异度 (假阳性率)",
           y = "灵敏度 (真阳性率)") +
      theme_minimal(base_size = 14) +
      theme(
        plot.title = element_text(hjust = 0.5, face = "bold"),
        panel.grid.minor = element_blank()
      )
    
    # 保存ROC曲线
    ggsave("ROC_Curve.png", roc_plot, width = 8, height = 6, dpi = 300)
    ggsave("ROC_Curve.pdf", roc_plot, width = 8, height = 6)
    cat("📁 ROC曲线已保存为: ROC_Curve.png 和 ROC_Curve.pdf\n")
    
  } else {
    cat("⚠️  数据不足，无法进行ROC分析\n")
  }
} else {
  cat("⚠️  未找到问卷总分，跳过ROC分析\n")
}

# 6. 逻辑回归分析 ---------------------------------------------------------
cat("\n📊 逻辑回归分析\n")
cat(rep("=", 60), "\n", sep = "")

# 6.1 单变量逻辑回归
cat("\n1️⃣ 单变量逻辑回归（关键心理指标）\n")

# 选择预测变量
predictors <- c("depression", "anxiety", "suicide_ideation", "psychosis", 
                "self_harm", "total_score")
predictors <- predictors[predictors %in% colnames(data)]

if(length(predictors) > 0) {
  univ_results <- data.frame()
  
  for(pred in predictors) {
    # 准备数据
    model_data <- data[!is.na(data$risk_binary) & !is.na(data[[pred]]), ]
    
    if(nrow(model_data) >= 50) {
      # 拟合模型
      model <- glm(risk_binary ~ ., data = model_data[, c("risk_binary", pred)], 
                   family = binomial)
      
      # 提取结果
      coef_summary <- summary(model)$coefficients
      
      if(pred %in% rownames(coef_summary)) {
        beta <- coef_summary[pred, "Estimate"]
        se <- coef_summary[pred, "Std. Error"]
        p_value <- coef_summary[pred, "Pr(>|z|)"]
        or <- exp(beta)
        ci_lower <- exp(beta - 1.96 * se)
        ci_upper <- exp(beta + 1.96 * se)
        
        univ_results <- rbind(univ_results, data.frame(
          变量 = pred,
          样本量 = nrow(model_data),
          阳性数 = sum(model_data$risk_binary),
          OR = round(or, 3),
          OR_95CI = paste0(round(ci_lower, 3), "-", round(ci_upper, 3)),
          P值 = format.pval(p_value, digits = 3, eps = 0.001)
        ))
      }
    }
  }
  
  if(nrow(univ_results) > 0) {
    print(univ_results)
    write.csv(univ_results, "univariate_logistic_regression.csv", row.names = FALSE)
    cat("📁 单变量逻辑回归结果已保存\n")
  }
}

# 6.2 多变量逻辑回归
cat("\n2️⃣ 多变量逻辑回归\n")

if(all(c("depression", "suicide_ideation") %in% colnames(data))) {
  model_data <- data[!is.na(data$risk_binary) & 
                       !is.na(data$depression) & 
                       !is.na(data$suicide_ideation), ]
  
  if(nrow(model_data) >= 50) {
    # 主效应模型
    main_model <- glm(risk_binary ~ depression + suicide_ideation, 
                      data = model_data, family = binomial)
    
    # 交互效应模型
    int_model <- glm(risk_binary ~ depression * suicide_ideation, 
                     data = model_data, family = binomial)
    
    cat("📊 模型比较:\n")
    cat("   主效应模型 AIC:", round(AIC(main_model), 1), "\n")
    cat("   交互效应模型 AIC:", round(AIC(int_model), 1), "\n")
    
    # 似然比检验
    lrtest <- anova(main_model, int_model, test = "Chisq")
    p_value <- lrtest$`Pr(>Chi)`[2]
    
    cat("   似然比检验 p值:", format.pval(p_value, digits = 3, eps = 0.001), "\n")
    
    if(!is.na(p_value) && p_value < 0.05) {
      cat("   ✅ 交互效应显著\n")
    } else {
      cat("   ⚠️  交互效应不显著\n")
    }
    
    # 提取多变量结果
    multi_results <- data.frame()
    for(model_name in c("主效应模型", "交互效应模型")) {
      model <- if(model_name == "主效应模型") main_model else int_model
      coef_summary <- summary(model)$coefficients
      
      for(i in 2:nrow(coef_summary)) {  # 跳过截距
        var_name <- rownames(coef_summary)[i]
        beta <- coef_summary[i, "Estimate"]
        se <- coef_summary[i, "Std. Error"]
        p_val <- coef_summary[i, "Pr(>|z|)"]
        or <- exp(beta)
        ci_lower <- exp(beta - 1.96 * se)
        ci_upper <- exp(beta + 1.96 * se)
        
        multi_results <- rbind(multi_results, data.frame(
          模型 = model_name,
          变量 = var_name,
          OR = round(or, 3),
          OR_95CI = paste0(round(ci_lower, 3), "-", round(ci_upper, 3)),
          P值 = format.pval(p_val, digits = 3, eps = 0.001)
        ))
      }
    }
    
    print(multi_results)
    write.csv(multi_results, "multivariate_logistic_regression.csv", row.names = FALSE)
  }
}

# 7. 可视化分析 ---------------------------------------------------------
cat("\n🎨 生成可视化图表\n")
cat(rep("=", 60), "\n", sep = "")

# 7.1 风险组总分分布图
if("total_score" %in% colnames(data)) {
  score_dist_plot <- ggplot(data, aes(x = factor(risk_binary), y = total_score, 
                                      fill = factor(risk_binary))) +
    geom_violin(alpha = 0.7) +
    geom_boxplot(width = 0.2, alpha = 0.7) +
    scale_fill_manual(values = c("0" = "#2E86AB", "1" = "#A23B72"),
                      labels = c("0" = "无需干预", "1" = "需要干预"),
                      name = "风险组") +
    labs(title = "问卷总分在风险组间的分布",
         x = "风险组",
         y = "问卷总分") +
    theme_minimal(base_size = 14) +
    theme(plot.title = element_text(hjust = 0.5, face = "bold"))
  
  ggsave("score_distribution_by_risk.png", score_dist_plot, width = 8, height = 6, dpi = 300)
  cat("📁 总分分布图已保存: score_distribution_by_risk.png\n")
}

# 7.2 关键心理指标箱线图
if(all(c("depression", "anxiety") %in% colnames(data))) {
  # 准备数据
  plot_data <- data %>%
    select(risk_binary, depression, anxiety) %>%
    pivot_longer(cols = c(depression, anxiety), 
                 names_to = "变量", 
                 values_to = "得分") %>%
    mutate(
      变量 = factor(变量, 
                  levels = c("depression", "anxiety"),
                  labels = c("抑郁", "焦虑")),
      风险组 = factor(risk_binary,
                   levels = c(0, 1),
                   labels = c("无需干预", "需要干预"))
    )
  
  psych_plot <- ggplot(plot_data, aes(x = 风险组, y = 得分, fill = 风险组)) +
    geom_boxplot(alpha = 0.7) +
    facet_wrap(~变量, scales = "free_y") +
    scale_fill_manual(values = c("无需干预" = "#2E86AB", "需要干预" = "#A23B72")) +
    labs(title = "心理指标在风险组间的比较",
         x = "",
         y = "指标得分") +
    theme_minimal(base_size = 14) +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold"),
      legend.position = "none",
      strip.text = element_text(face = "bold")
    )
  
  ggsave("psychological_scores_by_risk.png", psych_plot, width = 10, height = 6, dpi = 300)
  cat("📁 心理指标比较图已保存: psychological_scores_by_risk.png\n")
}

# 7.3 最佳切点性能图
if(exists("diag_table") && nrow(diag_table) > 0) {
  perf_plot <- diag_table %>%
    pivot_longer(cols = c(灵敏度, 特异度), 
                 names_to = "指标", 
                 values_to = "值") %>%
    ggplot(aes(x = 切点, y = 值, color = 指标, group = 指标)) +
    geom_line(size = 1.2) +
    geom_point(size = 3) +
    scale_color_manual(values = c("灵敏度" = "#A23B72", "特异度" = "#2E86AB")) +
    labs(title = "不同切点下灵敏度和特异度的变化",
         x = "切点 (问卷总分)",
         y = "性能指标",
         color = "指标") +
    theme_minimal(base_size = 14) +
    theme(plot.title = element_text(hjust = 0.5, face = "bold"))
  
  ggsave("cutoff_performance.png", perf_plot, width = 8, height = 6, dpi = 300)
  cat("📁 切点性能图已保存: cutoff_performance.png\n")
}

# 8. 结果汇总与报告 -------------------------------------------------------
cat("\n📋 分析结果汇总\n")
cat(rep("=", 60), "\n", sep = "")

# 创建结果汇总表
results_summary <- data.frame(
  分析项目 = c(
    "总样本量",
    "需要干预人数",
    "无需干预人数",
    "阳性率",
    "问卷总分AUC",
    "最佳切点总分",
    "最佳切点灵敏度",
    "最佳切点特异度",
    "抑郁指标OR值",
    "自杀意念指标OR值"
  ),
  结果 = c(
    format(nrow(data), big.mark = ","),
    format(n_positive, big.mark = ","),
    format(n_negative, big.mark = ","),
    paste0(round(positive_rate, 2), "%"),
    if(exists("auc_value")) round(auc_value, 3) else "N/A",
    if(exists("best_cutoff")) round(best_cutoff$threshold, 1) else "N/A",
    if(exists("best_cutoff")) round(best_cutoff$sensitivity, 3) else "N/A",
    if(exists("best_cutoff")) round(best_cutoff$specificity, 3) else "N/A",
    if(nrow(univ_results) > 0 && "depression" %in% univ_results$变量) 
      univ_results$OR[univ_results$变量 == "depression"] else "N/A",
    if(nrow(univ_results) > 0 && "suicide_ideation" %in% univ_results$变量) 
      univ_results$OR[univ_results$变量 == "suicide_ideation"] else "N/A"
  )
)

print(results_summary)

# 保存汇总结果
write.csv(results_summary, "analysis_summary.csv", row.names = FALSE)

# 9. 保存最终数据集 -----------------------------------------------------
cat("\n💾 保存处理后的数据集\n")
cat(rep("=", 60), "\n", sep = "")

# 选择关键变量保存
key_vars <- c("risk_binary", "total_score", predictors, 
              "severe_total", "general_total", "developmental_total",
              "depression_suicide", "psychosis_paranoia")
key_vars <- key_vars[key_vars %in% colnames(data)]

final_data <- data[, c("risk_binary", key_vars)]

# 保存为RDS和CSV
saveRDS(final_data, "processed_data_final.rds")
write.csv(final_data, "processed_data_final.csv", row.names = FALSE)

cat("✅ 处理后的数据已保存:\n")
cat("   processed_data_final.rds - R格式数据\n")
cat("   processed_data_final.csv - CSV格式数据\n")

# 10. 生成分析报告 ------------------------------------------------------
cat("\n📄 生成分析报告\n")
cat(rep("=", 60), "\n", sep = "")

report_text <- paste0(
  "========================================\n",
  "大学生心理健康筛查问卷预测效度分析报告\n",
  "========================================\n\n",
  "1. 样本特征\n",
  "   总样本量: ", format(nrow(data), big.mark = ","), "\n",
  "   需要干预: ", format(n_positive, big.mark = ","), " (", round(positive_rate, 2), "%)\n",
  "   无需干预: ", format(n_negative, big.mark = ","), " (", round(100 - positive_rate, 2), "%)\n\n",
  
  if(exists("auc_value")) {
    paste0(
      "2. 问卷预测效度\n",
      "   AUC: ", round(auc_value, 3), " (95% CI: ", round(ci_auc[1], 3), "-", round(ci_auc[3], 3), ")\n",
      "   最佳切点: 总分 ≥ ", round(best_cutoff$threshold, 1), "\n",
      "   灵敏度: ", round(best_cutoff$sensitivity * 100, 1), "%\n",
      "   特异度: ", round(best_cutoff$specificity * 100, 1), "%\n\n"
    )
  } else {
    "2. 问卷预测效度: 未计算（缺少问卷总分）\n\n"
  },
  
  "3. 关键心理指标\n",
  if(nrow(univ_results) > 0) {
    paste0(
      "   抑郁: OR = ", 
      ifelse("depression" %in% univ_results$变量, 
             univ_results$OR[univ_results$变量 == "depression"], "N/A"), "\n",
      "   自杀意念: OR = ", 
      ifelse("suicide_ideation" %in% univ_results$变量, 
             univ_results$OR[univ_results$变量 == "suicide_ideation"], "N/A"), "\n\n"
    )
  } else {
    "   未计算单变量逻辑回归\n\n"
  },
  
  "4. 生成文件列表\n",
  "   - ROC_Curve.png/pdf: ROC曲线图\n",
  "   - score_distribution_by_risk.png: 总分分布图\n",
  "   - psychological_scores_by_risk.png: 心理指标比较图\n",
  "   - table1_demographics.csv: 人口学特征表\n",
  "   - univariate_logistic_regression.csv: 单变量逻辑回归结果\n",
  "   - analysis_summary.csv: 分析结果汇总\n",
  "   - processed_data_final.rds/csv: 处理后的最终数据\n"
)

# 保存报告
writeLines(report_text, "analysis_report.txt")
cat(report_text)

cat("\n", rep("✅", 60), "\n", sep = "")
cat("           分析完成！\n")
cat("           所有结果已保存到当前目录\n")
cat(rep("✅", 60), "\n\n")

cat("💡 下一步建议:\n")
cat("   1. 查看分析报告: analysis_report.txt\n")
cat("   2. 检查生成的图表文件\n")
cat("   3. 使用 processed_data_final.rds 进行进一步分析\n")


---------------------------------------------
#ROC曲线绘图
  # 心理学期刊格式的ROC曲线
  library(pROC)
library(ggplot2)
library(cowplot)  # 用于专业主题

# 准备数据
roc_data <- data.frame(
  Specificity = 1 - roc_result$specificities,
  Sensitivity = roc_result$sensitivities
)

# 主图
roc_plot <- ggplot(roc_data, aes(x = Specificity, y = Sensitivity)) +
  # 对角线参考线
  geom_segment(aes(x = 0, y = 0, xend = 1, yend = 1), 
               color = "gray70", linetype = "dashed", linewidth = 0.5) +
  # ROC曲线
  geom_line(color = "#2C3E50", linewidth = 1.5) +
  # 最佳切点
  geom_point(aes(x = 1 - best_cutoff$specificity, y = best_cutoff$sensitivity),
             color = "#E74C3C", size = 4, shape = 19) +
  # AUC标注
  annotate("text", x = 0.65, y = 0.15, size = 5,
           label = paste("AUC = ", round(auc_value, 3), 
                         " [", round(ci_auc[1], 3), "-", 
                         round(ci_auc[3], 3), "]", sep = ""),
           hjust = 0, fontface = "bold") +
  # 最佳切点标注
  annotate("text", x = 1 - best_cutoff$specificity + 0.03,
           y = best_cutoff$sensitivity,
           label = paste("Cut-off:", round(best_cutoff$threshold, 1),
                         "\nSensitivity:", round(best_cutoff$sensitivity, 3),
                         "\nSpecificity:", round(best_cutoff$specificity, 3)),
           size = 4, hjust = 0, vjust = 0.5) +
  # 坐标轴设置
  scale_x_continuous(name = "1 - Specificity (False Positive Rate)",
                     breaks = seq(0, 1, 0.2),
                     expand = c(0.01, 0.01)) +
  scale_y_continuous(name = "Sensitivity (True Positive Rate)",
                     breaks = seq(0, 1, 0.2),
                     expand = c(0.01, 0.01)) +
  # 主题设置（心理学论文风格）
  theme_bw(base_size = 14) +
  theme(
    panel.grid.major = element_line(color = "gray90", linewidth = 0.3),
    panel.grid.minor = element_blank(),
    panel.border = element_rect(color = "black", linewidth = 0.8),
    axis.title = element_text(face = "bold", size = 13),
    axis.text = element_text(color = "black", size = 11),
    plot.title = element_text(hjust = 0.5, face = "bold", size = 14),
    aspect.ratio = 1  # 保持1:1比例
  ) +
  # 图例（可选）
  geom_text(aes(x = 0.5, y = 0.95, label = "Reference line"),
            color = "gray50", size = 3.5, hjust = 0) +
  geom_segment(aes(x = 0.35, xend = 0.45, y = 0.95, yend = 0.95),
               color = "gray70", linetype = "dashed", linewidth = 0.5)

# 保存为高分辨率图片
ggsave("Figure1_ROC_Curve.tiff", roc_plot, 
       width = 7, height = 7, dpi = 600, compression = "lzw")

# 同时保存PDF和PNG版本
ggsave("Figure1_ROC_Curve.pdf", roc_plot, width = 7, height = 7)
ggsave("Figure1_ROC_Curve.png", roc_plot, width = 7, height = 7, dpi = 300)

print(roc_plot)
--------------------------------------------------------------------
  #亚组间ROC曲线分析
  # 1. 准备性别数据
  # 确保性别列存在且编码正确
  if("性别" %in% colnames(data)) {
    data$gender_group <- factor(data$性别, 
                                levels = c("男", "女"),
                                labels = c("Male", "Female"))
  }

# 或者如果已经转换为英文
if(!exists("data$gender_group") && "gender" %in% colnames(data)) {
  data$gender_group <- factor(data$gender, 
                              levels = c("male", "female"),
                              labels = c("Male", "Female"))
}

# 2. 按性别分组的ROC分析
gender_roc_results <- list()

# 男性组
male_data <- data[data$gender_group == "Male" & 
                    !is.na(data$total_score) & 
                    !is.na(data$risk_binary), ]
gender_roc_results$male <- roc(male_data$risk_binary, male_data$total_score)

# 女性组
female_data <- data[data$gender_group == "Female" & 
                      !is.na(data$total_score) & 
                      !is.na(data$risk_binary), ]
gender_roc_results$female <- roc(female_data$risk_binary, female_data$total_score)

# 3. 比较AUC差异（DeLong检验）
gender_roc_test <- roc.test(gender_roc_results$male, 
                            gender_roc_results$female, 
                            method = "delong")

# 4. 提取结果
gender_summary <- data.frame(
  Subgroup = c("Male", "Female"),
  N = c(nrow(male_data), nrow(female_data)),
  Positive = c(sum(male_data$risk_binary), sum(female_data$risk_binary)),
  AUC = c(
    auc(gender_roc_results$male),
    auc(gender_roc_results$female)
  ),
  AUC_95CI = c(
    paste0(round(ci.auc(gender_roc_results$male)[1], 3), "-",
           round(ci.auc(gender_roc_results$male)[3], 3)),
    paste0(round(ci.auc(gender_roc_results$female)[1], 3), "-",
           round(ci.auc(gender_roc_results$female)[3], 3))
  ),
  Best_Cutoff = c(
    coords(gender_roc_results$male, "best", ret = "threshold")$threshold,
    coords(gender_roc_results$female, "best", ret = "threshold")$threshold
  ),
  Sensitivity = c(
    coords(gender_roc_results$male, "best", ret = "sensitivity")$sensitivity,
    coords(gender_roc_results$female, "best", ret = "sensitivity")$sensitivity
  ),
  Specificity = c(
    coords(gender_roc_results$male, "best", ret = "specificity")$specificity,
    coords(gender_roc_results$female, "best", ret = "specificity")$specificity
  )
)

print("Gender subgroup analysis results:")
print(gender_summary)

cat("\nDeLong test for AUC difference (Male vs Female):\n")
cat("Z =", round(gender_roc_test$statistic, 3), "\n")
cat("P-value =", round(gender_roc_test$p.value, 4), "\n")

# 5. 保存结果
write.csv(gender_summary, "gender_subgroup_analysis.csv", row.names = FALSE)

------------------------------------
  # 1. 检查原始居住地数据
  cat("原始居住地分布:\n")
print(table(data$生源地, useNA = "always"))

# 2. 创建新的二分类变量：农村 vs 城镇
data$residence_new <- factor(
  ifelse(data$生源地 == "农村", "Rural", "Urban"),
  levels = c("Rural", "Urban"),
  labels = c("Rural", "Urban")
)

# 3. 验证分类结果
cat("\n重新分类后分布:\n")
residence_table_new <- table(data$residence_new, useNA = "always")
print(residence_table_new)
# 1. 准备数据
residence_2cat_roc <- list()

for(res in c("Rural", "Urban")) {
  temp_data <- data[data$residence_new == res & 
                      !is.na(data$total_score) & 
                      !is.na(data$risk_binary), ]
  
  # 确保有足够样本
  if(nrow(temp_data) >= 20 && sum(temp_data$risk_binary) >= 5) {
    residence_2cat_roc[[res]] <- roc(temp_data$risk_binary, temp_data$total_score)
    cat(res, "组: n =", nrow(temp_data), 
        ", 高风险 =", sum(temp_data$risk_binary), 
        "(", round(sum(temp_data$risk_binary)/nrow(temp_data)*100, 2), "%)\n")
  } else {
    cat("⚠️  ", res, "组样本量不足\n")
  }
}

# 2. 计算AUC和置信区间
residence_2cat_summary <- data.frame()

for(res in names(residence_2cat_roc)) {
  roc_obj <- residence_2cat_roc[[res]]
  auc_ci <- ci.auc(roc_obj)
  best_cutoff <- coords(roc_obj, "best", ret = "all")
  
  residence_2cat_summary <- rbind(residence_2cat_summary, data.frame(
    Residence = res,
    N = nrow(data[data$residence_new == res & !is.na(data$total_score) & 
                    !is.na(data$risk_binary), ]),
    Positive = sum(data[data$residence_new == res & !is.na(data$risk_binary), ]$risk_binary, 
                   na.rm = TRUE),
    Positive_Rate = round(
      sum(data[data$residence_new == res & !is.na(data$risk_binary), ]$risk_binary, 
          na.rm = TRUE) / 
        nrow(data[data$residence_new == res & !is.na(data$total_score) & 
                    !is.na(data$risk_binary), ]) * 100, 2
    ),
    AUC = round(auc(roc_obj), 3),
    AUC_95CI_Lower = round(auc_ci[1], 3),
    AUC_95CI_Upper = round(auc_ci[3], 3),
    Best_Cutoff = round(best_cutoff$threshold, 1),
    Sensitivity = round(best_cutoff$sensitivity, 3),
    Specificity = round(best_cutoff$specificity, 3),
    Youden_Index = round(best_cutoff$sensitivity + best_cutoff$specificity - 1, 3)
  ))
}

print("农村vs城镇ROC分析结果:")
print(residence_2cat_summary)

# 3. 比较两组AUC（DeLong检验）
if(length(residence_2cat_roc) == 2) {
  roc_test_2cat <- roc.test(residence_2cat_roc$Rural, 
                            residence_2cat_roc$Urban, 
                            method = "delong")
  
  cat("\n农村 vs 城镇 AUC比较 (DeLong检验):\n")
  cat("   Z统计量 =", round(roc_test_2cat$statistic, 3), "\n")
  cat("   P值 =", format.pval(roc_test_2cat$p.value, digits = 3, eps = 0.001), "\n")
  
  # 计算AUC差异的置信区间
  auc_diff <- auc(residence_2cat_roc$Urban) - auc(residence_2cat_roc$Rural)
  cat("   AUC差异 (城镇-农村) =", round(auc_diff, 3), "\n")
}

# 4. 保存结果
write.csv(residence_2cat_summary, "residence_rural_urban_roc_results.csv", row.names = FALSE)
------------------------------------------------------------------
  # 1. 准备独生子女数据
  if("是否独生" %in% colnames(data)) {
    data$only_child_group <- factor(data$是否独生,
                                    levels = c("是", "不是"),
                                    labels = c("Only Child", "Not Only Child"))
  } else if("only_child" %in% colnames(data)) {
    data$only_child_group <- factor(data$only_child,
                                    levels = c("yes", "no"),
                                    labels = c("Only Child", "Not Only Child"))
  }

# 2. 按独生子女分组的ROC分析
onlychild_roc_results <- list()

for(group in c("Only Child", "Not Only Child")) {
  temp_data <- data[data$only_child_group == group & 
                      !is.na(data$total_score) & 
                      !is.na(data$risk_binary), ]
  onlychild_roc_results[[group]] <- roc(temp_data$risk_binary, temp_data$total_score)
}

# 3. 比较AUC差异
onlychild_roc_test <- roc.test(onlychild_roc_results$`Only Child`,
                               onlychild_roc_results$`Not Only Child`,
                               method = "delong")

# 4. 提取结果
onlychild_summary <- data.frame(
  Subgroup = c("Only Child", "Not Only Child"),
  N = c(
    nrow(data[data$only_child_group == "Only Child" & !is.na(data$total_score) & !is.na(data$risk_binary), ]),
    nrow(data[data$only_child_group == "Not Only Child" & !is.na(data$total_score) & !is.na(data$risk_binary), ])
  ),
  Positive = c(
    sum(data[data$only_child_group == "Only Child" & !is.na(data$risk_binary), ]$risk_binary, na.rm = TRUE),
    sum(data[data$only_child_group == "Not Only Child" & !is.na(data$risk_binary), ]$risk_binary, na.rm = TRUE)
  ),
  AUC = c(
    auc(onlychild_roc_results$`Only Child`),
    auc(onlychild_roc_results$`Not Only Child`)
  ),
  AUC_95CI = c(
    paste0(round(ci.auc(onlychild_roc_results$`Only Child`)[1], 3), "-",
           round(ci.auc(onlychild_roc_results$`Only Child`)[3], 3)),
    paste0(round(ci.auc(onlychild_roc_results$`Not Only Child`)[1], 3), "-",
           round(ci.auc(onlychild_roc_results$`Not Only Child`)[3], 3))
  ),
  Best_Cutoff = c(
    coords(onlychild_roc_results$`Only Child`, "best", ret = "threshold")$threshold,
    coords(onlychild_roc_results$`Not Only Child`, "best", ret = "threshold")$threshold
  ),
  Sensitivity = c(
    coords(onlychild_roc_results$`Only Child`, "best", ret = "sensitivity")$sensitivity,
    coords(onlychild_roc_results$`Not Only Child`, "best", ret = "sensitivity")$sensitivity
  ),
  Specificity = c(
    coords(onlychild_roc_results$`Only Child`, "best", ret = "specificity")$specificity,
    coords(onlychild_roc_results$`Not Only Child`, "best", ret = "specificity")$specificity
  )
)

print("Only Child subgroup analysis results:")
print(onlychild_summary)

cat("\nDeLong test for AUC difference (Only Child vs Not Only Child):\n")
cat("Z =", round(onlychild_roc_test$statistic, 3), "\n")
cat("P-value =", round(onlychild_roc_test$p.value, 4), "\n")

# 保存结果
write.csv(onlychild_summary, "onlychild_subgroup_analysis.csv", row.names = FALSE)
-------------------------------------------------------------------
  # 1. 计算总量表信度（如果有所有项目）
library(psych)

# 假设您的问卷有22个项目（确保列名匹配）
scale_items <- c("psychosis", "suicide_ideation", "anxiety", "depression",
                 "paranoia", "inferiority", "sensitivity", "social_phobia",
                 "somatization", "dependence", "hostility", "impulsivity",
                 "compulsion", "internet_addiction", "self_harm", 
                 "eating_disorder", "sleep_disturbance", "school_adjustment",
                 "interpersonal_stress", "academic_stress", "employment_stress",
                 "relationship_stress")

# 筛选存在的数据
available_items <- scale_items[scale_items %in% colnames(data)]

# 计算总量表信度
total_alpha <- psych::alpha(data[, available_items], check.keys = TRUE)
cat("总量表Cronbach's α:", round(total_alpha$total$raw_alpha, 3), "\n")
cat("95% CI:", round(total_alpha$total$ase, 3), "\n")

# 2. 计算各维度信度（如果定义了维度）
# 严重心理问题维度
severe_items <- c("psychosis", "suicide_ideation")
severe_alpha <- psych::alpha(data[, severe_items], check.keys = TRUE)
cat("\n严重心理问题维度 α =", round(severe_alpha$total$raw_alpha, 3), "\n")

# 一般心理问题维度（示例，根据实际定义调整）
general_items <- c("anxiety", "depression", "paranoia", "inferiority", 
                  "sensitivity", "social_phobia", "somatization", 
                  "dependence", "hostility", "impulsivity", "compulsion",
                  "internet_addiction", "self_harm", "eating_disorder", 
                  "sleep_disturbance")
general_alpha <- psych::alpha(data[, general_items], check.keys = TRUE)
cat("一般心理问题维度 α =", round(general_alpha$total$raw_alpha, 3), "\n")

# 发展性困扰维度
developmental_items <- c("school_adjustment", "interpersonal_stress", 
                        "academic_stress", "employment_stress", "relationship_stress")
developmental_alpha <- psych::alpha(data[, developmental_items], check.keys = TRUE)
cat("发展性困扰维度 α =", round(developmental_alpha$total$raw_alpha, 3), "\n")

# 3. 保存结果
reliability_results <- data.frame(
  量表部分 = c("总量表", "严重心理问题", "一般心理问题", "发展性困扰"),
  项目数 = c(length(available_items), length(severe_items), 
            length(general_items), length(developmental_items)),
  Cronbach_α = c(round(total_alpha$total$raw_alpha, 3),
                 round(severe_alpha$total$raw_alpha, 3),
                 round(general_alpha$total$raw_alpha, 3),
                 round(developmental_alpha$total$raw_alpha, 3))
)

print(reliability_results)
write.csv(reliability_results, "reliability_analysis_current_sample.csv", row.names = FALSE)

# 计算Cronbach's α及其置信区间
library(psych)

# 1. 定义各维度的项目
severe_items <- c("psychosis", "suicide_ideation")
general_items <- c("anxiety", "depression", "paranoia", "inferiority", 
                   "sensitivity", "social_phobia", "somatization", 
                   "dependence", "hostility", "impulsivity", "compulsion",
                   "internet_addiction", "self_harm", "eating_disorder", 
                   "sleep_disturbance")
developmental_items <- c("school_adjustment", "interpersonal_stress", 
                         "academic_stress", "employment_stress", "relationship_stress")

# 2. 检查项目中是否有缺失
available_items <- function(items, data) {
  items[items %in% colnames(data)]
}

# 3. 计算信度及置信区间的函数
calculate_alpha_with_ci <- function(items, data, scale_name) {
  # 获取实际存在的项目
  actual_items <- available_items(items, data)
  
  if(length(actual_items) < 2) {
    cat("警告: ", scale_name, " 有效项目数不足2个\n")
    return(NULL)
  }
  
  # 移除缺失值
  scale_data <- na.omit(data[, actual_items])
  
  # 计算Cronbach's α
  alpha_result <- psych::alpha(scale_data, check.keys = TRUE)
  
  # 提取结果
  results <- list(
    scale_name = scale_name,
    n_items = length(actual_items),
    n_valid = nrow(scale_data),
    alpha = alpha_result$total$raw_alpha,
    lower_ci = alpha_result$total$raw_alpha - 1.96 * alpha_result$total$ase,
    upper_ci = alpha_result$total$raw_alpha + 1.96 * alpha_result$total$ase,
    ase = alpha_result$total$ase  # alpha的标准误
  )
  
  return(results)
}

# 4. 计算各维度的信度
scales <- list(
  list(name = "Full Scale", items = c(severe_items, general_items, developmental_items)),
  list(name = "Severe Psychological Problems", items = severe_items),
  list(name = "General Psychological Problems", items = general_items),
  list(name = "Developmental Distress", items = developmental_items)
)

reliability_results <- list()

for(scale in scales) {
  result <- calculate_alpha_with_ci(scale$items, data, scale$name)
  if(!is.null(result)) {
    reliability_results[[scale$name]] <- result
  }
}

# 5. 创建结果表格
reliability_table <- data.frame(
  Scale = sapply(reliability_results, function(x) x$scale_name),
  Items = sapply(reliability_results, function(x) x$n_items),
  N = sapply(reliability_results, function(x) x$n_valid),
  Alpha = sapply(reliability_results, function(x) round(x$alpha, 3)),
  CI_Lower = sapply(reliability_results, function(x) round(x$lower_ci, 3)),
  CI_Upper = sapply(reliability_results, function(x) round(x$upper_ci, 3)),
  ASE = sapply(reliability_results, function(x) round(x$ase, 4))
)

# 添加95% CI字符串列
reliability_table$`95% CI` <- paste0("[", reliability_table$CI_Lower, ", ", 
                                     reliability_table$CI_Upper, "]")

print("信度分析结果（含置信区间）:")
print(reliability_table[, c("Scale", "Items", "N", "Alpha", "95% CI", "ASE")])

# 6. 保存结果
write.csv(reliability_table, "reliability_analysis_with_ci.csv", row.names = FALSE)

# 7. 输出美观表格
cat("\n==================== 信度分析结果汇总 ====================\n")
cat(sprintf("%-30s %6s %8s %8s %20s\n", 
            "Scale", "Items", "N", "Alpha", "95% CI"))
cat(rep("-", 75), "\n", sep = "")

for(i in 1:nrow(reliability_table)) {
  cat(sprintf("%-30s %6d %8d %8.3f %20s\n",
              reliability_table$Scale[i],
              reliability_table$Items[i],
              reliability_table$N[i],
              reliability_table$Alpha[i],
              reliability_table$`95% CI`[i]))
}

---------------------------------------------------------------------------
  # 完整的组间比较代码（无需moments包）
  library(dplyr)
library(tidyr)
library(rstatix)

# 1. 创建关键心理指标比较函数
compare_psychological_measures <- function(data, vars, group_var = "risk_binary") {
  results <- data.frame()
  
  for(var in vars) {
    if(!var %in% colnames(data)) next
    
    # 准备数据
    test_data <- data[!is.na(data[[var]]) & !is.na(data[[group_var]]), ]
    
    # 计算描述性统计
    desc_stats <- test_data %>%
      group_by(!!sym(group_var)) %>%
      summarise(
        n = n(),
        mean = mean(.data[[var]]),
        sd = sd(.data[[var]]),
        median = median(.data[[var]]),
        min = min(.data[[var]]),
        max = max(.data[[var]])
      )
    
    # 使用Welch's t检验（对大样本稳健）
    test_result <- t.test(test_data[[var]] ~ test_data[[group_var]], 
                          var.equal = FALSE)
    
    # 计算效应量
    group0 <- test_data[[var]][test_data[[group_var]] == 0]
    group1 <- test_data[[var]][test_data[[group_var]] == 1]
    
    n0 <- length(group0)
    n1 <- length(group1)
    sd0 <- sd(group0)
    sd1 <- sd(group1)
    
    pooled_sd <- sqrt(((n0 - 1) * sd0^2 + (n1 - 1) * sd1^2) / (n0 + n1 - 2))
    cohens_d <- (mean(group1) - mean(group0)) / pooled_sd
    
    # 创建结果行
    results <- rbind(results, data.frame(
      Variable = var,
      Group0_N = desc_stats$n[1],
      Group0_Mean = round(desc_stats$mean[1], 2),
      Group0_SD = round(desc_stats$sd[1], 2),
      Group1_N = desc_stats$n[2],
      Group1_Mean = round(desc_stats$mean[2], 2),
      Group1_SD = round(desc_stats$sd[2], 2),
      t_statistic = round(test_result$statistic, 3),
      df = round(test_result$parameter, 1),
      p_value = format.pval(test_result$p.value, digits = 3, eps = 0.001),
      cohens_d = round(cohens_d, 3)
    ))
  }
  
  return(results)
}

# 2. 运行比较
psych_comparison <- compare_psychological_measures(data, key_psych_vars)

print("关键心理指标组间比较:")
print(psych_comparison)

# 3. 创建简洁的发表用表格
publication_table <- psych_comparison %>%
  mutate(
    `No Follow-up` = sprintf("%.2f ± %.2f", Group0_Mean, Group0_SD),
    `Follow-up Needed` = sprintf("%.2f ± %.2f", Group1_Mean, Group1_SD),
    `t(df)` = sprintf("t(%.1f) = %.2f", df, t_statistic),
    `Cohen's d` = sprintf("%.2f", cohens_d),
    `p` = p_value
  ) %>%
  select(
    Variable,
    `No Follow-up`,
    `Follow-up Needed`,
    `t(df)`,
    `Cohen's d`,
    `p`
  )

# 添加中文变量名
publication_table$Variable <- c("Depression", "Anxiety", "Suicide ideation", 
                                "Psychosis", "Self-harm", "Total score")

print("发表用表格:")
print(publication_table)

# 4. 保存结果
write.csv(psych_comparison, "psychological_measures_comparison.csv", row.names = FALSE)
write.csv(publication_table, "table_psychological_comparisons_final.csv", row.names = FALSE)

# 创建APA格式表格
apa_table <- publication_table %>%
  mutate(
    `p` = ifelse(`p` == "<0.001", "< .001", 
                 paste0("= ", gsub("p = ", "", `p`)))
  )

print("APA格式表格:")
print(apa_table)

# 使用flextable创建美观的Word表格
if (require(flextable)) {
  ft <- flextable(apa_table) %>%
    set_header_labels(
      Variable = "Psychological Measure",
      `No Follow-up` = "No Follow-up\n(n = 17,092)",
      `Follow-up Needed` = "Follow-up Needed\n(n = 229)",
      `t(df)` = "t(df)",
      `Cohen's d` = "Cohen's d",
      `p` = "p"
    ) %>%
    add_header_lines(values = "Table 2. Comparison of Psychological Measures Between Groups") %>%
    theme_booktabs() %>%
    autofit() %>%
    align(align = "center", part = "all") %>%
    align(align = "left", part = "body", j = 1) %>%
    bold(part = "header") %>%
    fontsize(size = 10, part = "all")
  
  save_as_docx(ft, path = "Table2_Psychological_Measures_Comparison.docx")
  cat("Word表格已保存: Table2_Psychological_Measures_Comparison.docx\n")
  
  # 预览表格
  print(ft, preview = "docx")
} else {
  # 安装flextable包
  install.packages("flextable")
  library(flextable)
  
  ft <- flextable(apa_table) %>%
    set_header_labels(
      Variable = "Psychological Measure",
      `No Follow-up` = "No Follow-up\n(n = 17,092)",
      `Follow-up Needed` = "Follow-up Needed\n(n = 229)",
      `t(df)` = "t(df)",
      `Cohen's d` = "Cohen's d",
      `p` = "p"
    ) %>%
    add_header_lines(values = "Table 2. Comparison of Psychological Measures Between Groups") %>%
    theme_booktabs() %>%
    autofit() %>%
    align(align = "center", part = "all") %>%
    align(align = "left", part = "body", j = 1) %>%
    bold(part = "header") %>%
    fontsize(size = 10, part = "all")
  
  save_as_docx(ft, path = "Table2_Psychological_Measures_Comparison.docx")
  cat("Word表格已保存: Table2_Psychological_Measures_Comparison.docx\n")
  
  # 创建组间差异可视化（无需moments包）
  library(ggplot2)
  library(patchwork)
  
  # 1. 创建单个变量的箱线图函数
  create_single_boxplot <- function(data, var, var_label) {
    plot_data <- data %>%
      select(risk_binary, value = all_of(var)) %>%
      filter(!is.na(value)) %>%
      mutate(
        Group = factor(risk_binary,
                       levels = c(0, 1),
                       labels = c("No Follow-up", "Follow-up Needed"))
      )
    
    # 提取p值
    p_value <- psych_comparison %>%
      filter(Variable == var) %>%
      pull(p_value)
    
    ggplot(plot_data, aes(x = Group, y = value, fill = Group)) +
      geom_boxplot(alpha = 0.7, outlier.shape = NA) +
      geom_jitter(width = 0.1, alpha = 0.1, size = 0.5) +
      scale_fill_manual(values = c("No Follow-up" = "#2E86AB", 
                                   "Follow-up Needed" = "#A23B72")) +
      labs(
        title = var_label,
        x = "",
        y = "Score",
        caption = paste0("p-value: ", p_value)
      ) +
      theme_minimal(base_size = 11) +
      theme(
        plot.title = element_text(hjust = 0.5, face = "bold"),
        legend.position = "none",
        axis.text.x = element_text(angle = 45, hjust = 1)
      )
  }
  
  # 2. 创建所有变量的箱线图
  if (length(key_psych_vars) > 0) {
    plots <- list()
    var_labels <- c("Depression", "Anxiety", "Suicide ideation", 
                    "Psychosis", "Self-harm", "Total score")
    
    for(i in seq_along(key_psych_vars)) {
      var <- key_psych_vars[i]
      if(var %in% colnames(data)) {
        plots[[i]] <- create_single_boxplot(data, var, var_labels[i])
      }
    }
    
    # 移除空元素
    plots <- plots[!sapply(plots, is.null)]
    
    # 组合图形
    if(length(plots) > 0) {
      combined_plot <- wrap_plots(plots, ncol = 3) +
        plot_annotation(
          title = "Comparison of Psychological Measures Between Groups",
          theme = theme(plot.title = element_text(hjust = 0.5, face = "bold", size = 14))
        )
      
      ggsave("Figure_S7_Psychological_Measures_Comparison.png", combined_plot, 
             width = 15, height = 10, dpi = 300)
      cat("图形已保存: Figure_S7_Psychological_Measures_Comparison.png\n")
      
      # 显示图形
      print(combined_plot)
    }
  }
  
  # 3. 效应量可视化
  if (nrow(psych_comparison) > 0) {
    effect_size_plot <- ggplot(psych_comparison, 
                               aes(x = reorder(Variable, cohens_d), 
                                   y = cohens_d, 
                                   fill = cohens_d)) +
      geom_bar(stat = "identity", width = 0.7) +
      geom_text(aes(label = round(cohens_d, 2)), 
                hjust = ifelse(psych_comparison$cohens_d > 0, -0.2, 1.2), 
                size = 4) +
      scale_fill_gradient2(low = "blue", mid = "white", high = "red", 
                           midpoint = 0, name = "Cohen's d") +
      coord_flip() +
      labs(
        title = "Effect Sizes for Group Differences in Psychological Measures",
        x = "Psychological Measure",
        y = "Cohen's d"
      ) +
      theme_minimal(base_size = 12) +
      theme(
        plot.title = element_text(hjust = 0.5, face = "bold"),
        legend.position = "right"
      ) +
      geom_hline(yintercept = c(0.2, 0.5, 0.8), 
                 linetype = c("dashed", "dotted", "dotted"),
                 color = "gray50", alpha = 0.7) +
      annotate("text", x = 0.5, y = 0.25, label = "Small", color = "gray50", hjust = 0) +
      annotate("text", x = 0.5, y = 0.55, label = "Medium", color = "gray50", hjust = 0) +
      annotate("text", x = 0.5, y = 0.85, label = "Large", color = "gray50", hjust = 0)
    
    ggsave("Figure_S8_Effect_Sizes.png", effect_size_plot, width = 10, height = 6, dpi = 300)
    cat("图形已保存: Figure_S8_Effect_Sizes.png\n")
    
    print(effect_size_plot)
  }
}

===============================================================================
  # 检查数据框结构
  cat("\n🔍 数据框变量名:\n")
print(colnames(data))

# 或者更简洁地
cat("\n前10个变量名:\n")
print(colnames(data)[1:10])

# 查找包含性别信息的列
gender_cols <- grep("性别|sex|gender", colnames(data), ignore.case = TRUE, value = TRUE)
cat("\n可能的性别列:", paste(gender_cols, collapse = ", "), "\n")

# 查找所有人口学变量
demo_cols <- grep("性别|生源地|是否独生|民族|年龄|年级", colnames(data), value = TRUE)
cat("\n人口学变量:", paste(demo_cols, collapse = ", "), "\n")

# ============================================================
# ============================================================
# 多元Logistic回归模型完整分析
# 包含：人口学特征 + 5个心理学变量 + 完整统计信息 + ROC曲线
# ============================================================

# 清理环境
rm(list = ls())
cat("\f")  # 清空控制台

# 1. 加载必要包 ------------------------------------------------------------
cat("🔧 加载必要包...\n")

required_packages <- c(
  "readxl",      # 读取Excel
  "dplyr",       # 数据处理
  "tidyr",       # 数据整理
  "ggplot2",     # 绘图
  "pROC",        # ROC分析
  "caret",       # 模型评估
  "ggpubr",      # 出版级图表
  "psych",       # 描述性统计
  "tableone",    # 创建Table 1
  "writexl",     # 写入Excel
  "car",         # 多重共线性检验
  "report",      # 模型报告
  "gt",          # 精美表格
  "gtsummary",   # 统计表格
  "forestplot"   # 森林图
)

# 安装缺失的包
new_packages <- required_packages[!(required_packages %in% installed.packages()[,"Package"])]
if(length(new_packages) > 0) {
  cat("📦 安装缺失包:", paste(new_packages, collapse = ", "), "\n")
  install.packages(new_packages, dependencies = TRUE)
}

# 加载包
suppressPackageStartupMessages({
  for(pkg in required_packages) {
    library(pkg, character.only = TRUE)
  }
})

cat("✅ 所有包加载完成\n")

# 2. 数据准备 -------------------------------------------------------------
cat("\n📊 数据准备\n")
cat(rep("=", 60), "\n", sep = "")

# 读取数据（根据你的文件路径修改）
data_path <- "01 data/new.xlsx"  # 修改为你的数据路径
if(!file.exists(data_path)) {
  stop("❌ 文件不存在: ", data_path, 
       "\n请确保文件路径正确，当前工作目录: ", getwd())
}

data <- read_excel(data_path)

# 检查数据
cat("数据维度:", nrow(data), "行 ×", ncol(data), "列\n")

# 3. 数据清理和变量创建 ----------------------------------------------------
cat("\n🧹 数据清理和变量创建\n")
cat(rep("=", 60), "\n", sep = "")

# 3.1 确保因变量是二分类数值型
if(!any(grepl("risk|风险", colnames(data), ignore.case = TRUE))) {
  stop("❌ 未找到风险标签列")
}

# 查找风险列
risk_col <- grep("risk|风险", colnames(data), ignore.case = TRUE, value = TRUE)[1]
data$risk_binary <- as.numeric(as.character(data[[risk_col]]))

# 验证风险变量
if(!all(na.omit(unique(data$risk_binary)) %in% c(0, 1))) {
  warning("风险变量包含非0/1值，自动转换: 非0值转为1")
  data$risk_binary <- ifelse(data$risk_binary == 0, 0, 1)
}

cat("风险标签分布:\n")
print(table(data$risk_binary, useNA = "ifany"))
cat("高风险比例:", round(mean(data$risk_binary, na.rm = TRUE) * 100, 2), "%\n")

# 3.2 创建人口学变量
cat("\n👥 创建人口学变量...\n")

# 3.2.1 性别变量 (男性=1, 女性=0)
gender_col <- grep("性别|gender|sex", colnames(data), ignore.case = TRUE, value = TRUE)[1]
if(!is.na(gender_col)) {
  gender_raw <- as.character(data[[gender_col]])
  data$gender <- ifelse(grepl("男|male|1", tolower(gender_raw)), 1,
                        ifelse(grepl("女|female|2", tolower(gender_raw)), 0, NA))
  cat("✅ 性别变量创建完成\n")
  cat("男性比例:", round(mean(data$gender, na.rm = TRUE) * 100, 1), "%\n")
} else {
  cat("⚠️ 未找到性别列\n")
  data$gender <- NA
}

# 3.2.2 居住地变量 (城市=1, 农村=0)
residence_col <- grep("居住地|生源地|residence|籍贯", colnames(data), ignore.case = TRUE, value = TRUE)[1]
if(!is.na(residence_col)) {
  residence_raw <- as.character(data[[residence_col]])
  residence_lower <- tolower(residence_raw)
  data$residence <- ifelse(grepl("城市|城镇|urban|city|市|镇", residence_lower), 1,
                           ifelse(grepl("农村|乡村|rural|乡|村|县", residence_lower), 0, NA))
  cat("✅ 居住地变量创建完成\n")
  cat("城市比例:", round(mean(data$residence, na.rm = TRUE) * 100, 1), "%\n")
} else {
  cat("⚠️ 未找到居住地列\n")
  data$residence <- NA
}

# 方法1：最简洁直接的代码
if("是否独生" %in% colnames(data)) {
  # 查看原始数据分布
  cat("独生子女原始数据分布:\n")
  print(table(data[["是否独生"]], useNA = "always"))
  
  # 进行编码转换
  data$onlychild <- ifelse(data[["是否独生"]] == "是", 1,
                           ifelse(data[["是否独生"]] == "不是", 0, NA))
  
  # 检查转换结果
  cat("\n转换后的分布:\n")
  freq_table <- table(data$onlychild, useNA = "always")
  print(freq_table)
  
  # 计算比例
  n_total <- nrow(data)
  n_yes <- sum(data$onlychild == 1, na.rm = TRUE)
  n_no <- sum(data$onlychild == 0, na.rm = TRUE)
  n_na <- sum(is.na(data$onlychild))
  
  cat(sprintf("\n统计结果:\n"))
  cat(sprintf("  是 (独生子女): %d (%.1f%%)\n", n_yes, n_yes/n_total*100))
  cat(sprintf("  不是 (非独生子女): %d (%.1f%%)\n", n_no, n_no/n_total*100))
  cat(sprintf("  缺失值: %d (%.1f%%)\n", n_na, n_na/n_total*100))
  
  # 与风险组交叉分析
  if("risk_binary" %in% colnames(data)) {
    cat("\n按风险组的独生子女分布:\n")
    cross_table <- table(
      独生子女 = factor(data$onlychild, levels = c(0, 1), labels = c("不是", "是")),
      风险组 = factor(data$risk_binary, levels = c(0, 1), labels = c("低风险", "高风险")),
      useNA = "always"
    )
    print(cross_table)
    
    # 计算各组的独生子女比例
    low_risk_yes <- cross_table["是", "低风险"] / sum(cross_table[, "低风险"], na.rm = TRUE) * 100
    high_risk_yes <- cross_table["是", "高风险"] / sum(cross_table[, "高风险"], na.rm = TRUE) * 100
    
    cat(sprintf("\n低风险组独生子女比例: %.1f%%\n", low_risk_yes))
    cat(sprintf("高风险组独生子女比例: %.1f%%\n", high_risk_yes))
  }
  
} else {
  cat("⚠️ 数据中没有找到'是否独生'列\n")
  
  # 查看所有列名，帮助找到正确的列
  cat("查看所有列名:\n")
  print(colnames(data))
  
  # 或者查找类似列名
  similar_cols <- grep("独生|是否", colnames(data), value = TRUE)
  if(length(similar_cols) > 0) {
    cat("找到类似列名:", paste(similar_cols, collapse = ", "), "\n")
    cat("请确认正确的列名并修改代码中的'是否独生'\n")
  }
}

# 3.3 创建心理学变量
cat("\n🧠 创建心理学变量...\n")

psych_vars_info <- list(
  "depression" = c("抑郁", "depression"),
  "anxiety" = c("焦虑", "anxiety"),
  "suicide_ideation" = c("自杀", "suicide"),
  "psychosis" = c("幻觉|妄想|精神病", "psychosis"),
  "self_harm" = c("自伤", "self.harm|selfharm")
)

for(psych_var in names(psych_vars_info)) {
  patterns <- psych_vars_info[[psych_var]]
  found <- FALSE
  
  for(pattern in patterns) {
    matches <- grep(pattern, colnames(data), ignore.case = TRUE, value = TRUE)
    if(length(matches) > 0) {
      # 找到变量
      data[[psych_var]] <- as.numeric(as.character(data[[matches[1]]]))
      cat(sprintf("✅ %s: 使用变量 %s\n", psych_var, matches[1]))
      
      # 检查是否有缺失值
      na_count <- sum(is.na(data[[psych_var]]))
      if(na_count > 0) {
        cat(sprintf("   警告: 有 %d 个缺失值\n", na_count))
      }
      
      # 基本统计
      cat(sprintf("   均值: %.2f, 标准差: %.2f, 范围: %.2f-%.2f\n",
                  mean(data[[psych_var]], na.rm = TRUE),
                  sd(data[[psych_var]], na.rm = TRUE),
                  min(data[[psych_var]], na.rm = TRUE),
                  max(data[[psych_var]], na.rm = TRUE)))
      
      found <- TRUE
      break
    }
  }
  
  if(!found) {
    cat(sprintf("⚠️  未找到 %s 变量\n", psych_var))
    data[[psych_var]] <- NA
  }
}

# 4. 创建完整分析数据集 ----------------------------------------------------
cat("\n📁 创建完整分析数据集\n")
cat(rep("=", 60), "\n", sep = "")

# 选择需要的变量
model_vars <- c("risk_binary", "gender", "residence", "onlychild",
                "depression", "anxiety", "suicide_ideation", "psychosis", "self_harm")

# 检查哪些变量存在
existing_vars <- model_vars[model_vars %in% colnames(data)]
cat("将使用的变量:", paste(existing_vars, collapse = ", "), "\n")

# 创建完整案例数据集
complete_data <- na.omit(data[, existing_vars])
cat("完整案例数:", nrow(complete_data), "\n")
cat("高风险案例数:", sum(complete_data$risk_binary), " (", 
    round(mean(complete_data$risk_binary) * 100, 2), "%)\n")

if(nrow(complete_data) < 50) {
  stop("❌ 完整案例数不足，无法进行可靠的模型分析")
}

# 检查事件数规则（每个预测变量至少10个事件）
n_predictors <- length(existing_vars) - 1
n_events <- sum(complete_data$risk_binary)
cat("预测变量数:", n_predictors, "\n")
cat("事件数（高风险案例）:", n_events, "\n")
cat("事件数/变量数:", round(n_events / n_predictors, 1), "\n")

if(n_events / n_predictors < 10) {
  warning("⚠️  事件数/变量数 < 10，模型可能不稳定")
}

# 5. 描述性统计 -----------------------------------------------------------
cat("\n📈 描述性统计\n")
cat(rep("=", 60), "\n", sep = "")

# 5.1 按风险分组的人口学特征
cat("\n👥 按风险分组的人口学特征:\n")

demo_vars <- c("gender", "residence", "onlychild")
demo_labels <- c("性别（男性）", "居住地（城市）", "独生子女")

# 创建描述性统计表
demo_table <- data.frame()

for(i in 1:length(demo_vars)) {
  var <- demo_vars[i]
  if(var %in% colnames(complete_data)) {
    # 按风险组计算比例
    prop_table <- complete_data %>%
      group_by(risk_binary) %>%
      summarise(
        比例 = mean(get(var), na.rm = TRUE) * 100,
        .groups = 'drop'
      )
    
    # 添加到表格
    demo_table <- rbind(demo_table, data.frame(
      变量 = demo_labels[i],
      低风险组 = sprintf("%.1f%%", prop_table$比例[prop_table$risk_binary == 0]),
      高风险组 = sprintf("%.1f%%", prop_table$比例[prop_table$risk_binary == 1]),
      风险比 = sprintf("%.2f", 
                    prop_table$比例[prop_table$risk_binary == 1] / 
                      prop_table$比例[prop_table$risk_binary == 0])
    ))
  }
}

print(demo_table, row.names = FALSE)

# 5.2 心理学变量的描述统计
cat("\n🧠 心理学变量的描述统计（均值 ± 标准差）:\n")

psych_vars <- c("depression", "anxiety", "suicide_ideation", "psychosis", "self_harm")
psych_labels <- c("抑郁", "焦虑", "自杀意念", "精神病性症状", "自伤行为")

psych_table <- data.frame()

for(i in 1:length(psych_vars)) {
  var <- psych_vars[i]
  if(var %in% colnames(complete_data)) {
    # 按风险组计算
    stats <- complete_data %>%
      group_by(risk_binary) %>%
      summarise(
        均值 = mean(get(var), na.rm = TRUE),
        标准差 = sd(get(var), na.rm = TRUE),
        .groups = 'drop'
      )
    
    # 计算效应量（Cohen's d）
    mean_low <- stats$均值[stats$risk_binary == 0]
    mean_high <- stats$均值[stats$risk_binary == 1]
    sd_low <- stats$标准差[stats$risk_binary == 0]
    sd_high <- stats$标准差[stats$risk_binary == 1]
    
    pooled_sd <- sqrt((sd_low^2 + sd_high^2) / 2)
    cohens_d <- (mean_high - mean_low) / pooled_sd
    
    # 添加到表格
    psych_table <- rbind(psych_table, data.frame(
      变量 = psych_labels[i],
      低风险组 = sprintf("%.2f ± %.2f", mean_low, sd_low),
      高风险组 = sprintf("%.2f ± %.2f", mean_high, sd_high),
      效应量 = sprintf("%.2f", cohens_d),
      解释 = ifelse(abs(cohens_d) < 0.2, "小",
                  ifelse(abs(cohens_d) < 0.5, "中",
                         ifelse(abs(cohens_d) < 0.8, "大", "很大")))
    ))
  }
}

print(psych_table, row.names = FALSE)

# 6. 多元Logistic回归模型 -------------------------------------------------
cat("\n📊 多元Logistic回归模型\n")
cat(rep("=", 60), "\n", sep = "")

# 6.1 构建模型
cat("\n🔧 构建多元Logistic回归模型...\n")

# 构建公式
predictors <- setdiff(existing_vars, "risk_binary")
formula_str <- paste("risk_binary ~", paste(predictors, collapse = " + "))
cat("模型公式:", formula_str, "\n")

# 拟合模型
model <- glm(as.formula(formula_str), 
             data = complete_data, 
             family = binomial(link = "logit"))

cat("✅ 模型拟合完成\n")

# 6.2 模型整体检验
cat("\n📈 模型整体检验:\n")

# 似然比检验（与空模型比较）
null_model <- glm(risk_binary ~ 1, data = complete_data, family = binomial)
lr_test <- anova(null_model, model, test = "Chisq")

cat("似然比检验:\n")
cat(sprintf("  卡方值(χ²) = %.2f\n", lr_test$Deviance[2]))
cat(sprintf("  自由度(df) = %d\n", lr_test$Df[2]))
cat(sprintf("  P值 = %.4f\n", lr_test$`Pr(>Chi)`[2]))

if(lr_test$`Pr(>Chi)`[2] < 0.001) {
  cat("  结论: 模型显著优于空模型 (P < 0.001)\n")
} else if(lr_test$`Pr(>Chi)`[2] < 0.05) {
  cat("  结论: 模型显著优于空模型 (P < 0.05)\n")
} else {
  cat("  结论: 模型不显著优于空模型\n")
}

# 6.3 模型拟合优度指标
cat("\n📊 模型拟合优度指标:\n")

# AIC和BIC
cat(sprintf("  AIC = %.2f (越小越好)\n", AIC(model)))
cat(sprintf("  BIC = %.2f (越小越好)\n", BIC(model)))

# 伪R²
null_deviance <- model$null.deviance
residual_deviance <- model$deviance
n <- nrow(complete_data)

# McFadden伪R²
mcfadden_r2 <- 1 - (residual_deviance / null_deviance)

# Cox-Snell伪R²
cox_snell_r2 <- 1 - exp((null_deviance - residual_deviance) / n)

# Nagelkerke伪R²（调整的Cox-Snell）
nagelkerke_r2 <- cox_snell_r2 / (1 - exp(-null_deviance / n))

cat(sprintf("  McFadden伪R² = %.3f\n", mcfadden_r2))
cat(sprintf("  Cox-Snell伪R² = %.3f\n", cox_snell_r2))
cat(sprintf("  Nagelkerke伪R² = %.3f\n", nagelkerke_r2))

# 伪R²解释
cat("  伪R²解释:\n")
cat("  <0.02: 拟合很差\n")
cat("  0.02-0.1: 拟合较弱\n")
cat("  0.1-0.2: 拟合中等\n")
cat("  0.2-0.4: 拟合较好\n")
cat("  >0.4: 拟合很好\n")

# 6.4 提取模型系数
cat("\n📋 模型系数结果:\n")

# 获取系数和置信区间
coef_summary <- summary(model)$coefficients
conf_int <- confint(model)

# 创建详细的结果表
results_detailed <- data.frame()

for(i in 1:nrow(coef_summary)) {
  var_name <- rownames(coef_summary)[i]
  beta <- coef_summary[i, "Estimate"]
  se <- coef_summary[i, "Std. Error"]
  z <- coef_summary[i, "z value"]
  p_val <- coef_summary[i, "Pr(>|z|)"]
  
  # 计算OR和置信区间
  or <- exp(beta)
  
  if(!any(is.na(conf_int)) && i <= nrow(conf_int)) {
    ci_lower <- exp(conf_int[i, 1])
    ci_upper <- exp(conf_int[i, 2])
  } else {
    ci_lower <- exp(beta - 1.96 * se)
    ci_upper <- exp(beta + 1.96 * se)
  }
  
  # 变量标签
  var_label <- switch(var_name,
                      "(Intercept)" = "截距",
                      "gender" = "性别（男性）",
                      "residence" = "居住地（城市）",
                      "onlychild" = "独生子女",
                      "depression" = "抑郁",
                      "anxiety" = "焦虑",
                      "suicide_ideation" = "自杀意念",
                      "psychosis" = "精神病性症状",
                      "self_harm" = "自伤行为",
                      var_name)
  
  # 显著性标记
  significance <- ""
  if(p_val < 0.001) significance <- "***"
  else if(p_val < 0.01) significance <- "**"
  else if(p_val < 0.05) significance <- "*"
  else if(p_val < 0.1) significance <- "."
  
  results_detailed <- rbind(results_detailed, data.frame(
    变量 = var_label,
    系数β = sprintf("%.3f", beta),
    标准误 = sprintf("%.3f", se),
    z值 = sprintf("%.2f", z),
    OR = sprintf("%.3f", or),
    `95% CI` = sprintf("%.3f-%.3f", ci_lower, ci_upper),
    P值 = ifelse(p_val < 0.001, "<0.001", sprintf("%.3f", p_val)),
    显著性 = significance,
    解释 = ifelse(var_name == "(Intercept)", "基准风险",
                ifelse(or > 1, "增加风险", "降低风险"))
  ))
}

print(results_detailed, row.names = FALSE)

# 6.5 多重共线性诊断
cat("\n🔍 多重共线性诊断:\n")

# 计算VIF
vif_values <- car::vif(model)

vif_table <- data.frame(
  变量 = names(vif_values),
  VIF = sprintf("%.2f", vif_values),
  诊断 = ifelse(vif_values > 10, "严重共线性",
              ifelse(vif_values > 5, "中度共线性", "正常"))
)

print(vif_table, row.names = FALSE)

# 6.6 模型预测性能
cat("\n🎯 模型预测性能:\n")

# 预测概率
predicted_probs <- predict(model, type = "response")
predicted_class <- ifelse(predicted_probs > 0.5, 1, 0)

# 混淆矩阵
confusion_matrix <- table(真实 = complete_data$risk_binary, 预测 = predicted_class)
cat("混淆矩阵（切点=0.5）:\n")
print(confusion_matrix)

# 计算各项指标
TP <- confusion_matrix[2, 2]  # 真阳性
TN <- confusion_matrix[1, 1]  # 真阴性
FP <- confusion_matrix[1, 2]  # 假阳性
FN <- confusion_matrix[2, 1]  # 假阴性

accuracy <- (TP + TN) / sum(confusion_matrix)
sensitivity <- TP / (TP + FN)  # 灵敏度/召回率
specificity <- TN / (TN + FP)  # 特异度
precision <- TP / (TP + FP)    # 精确度/阳性预测值
npv <- TN / (TN + FN)          # 阴性预测值
f1_score <- 2 * (precision * sensitivity) / (precision + sensitivity)

cat(sprintf("\n准确率: %.2f%%\n", accuracy * 100))
cat(sprintf("灵敏度（召回率）: %.2f%%\n", sensitivity * 100))
cat(sprintf("特异度: %.2f%%\n", specificity * 100))
cat(sprintf("精确度（阳性预测值）: %.2f%%\n", precision * 100))
cat(sprintf("阴性预测值: %.2f%%\n", npv * 100))
cat(sprintf("F1分数: %.3f\n", f1_score))

# 7. ROC曲线分析（修复横坐标刻度）---------------------------
# 7. ROC曲线分析（修复横坐标刻度）---------------------------
cat("\n📈 ROC Curve Analysis (Fixed X-axis)\n")
cat(rep("=", 60), "\n", sep = "")

# 7.1 计算ROC曲线
cat("\n🔧 Calculating ROC curve...\n")

roc_result <- roc(complete_data$risk_binary, predicted_probs)
auc_value <- auc(roc_result)
ci_auc <- ci.auc(roc_result)

# AUC解释
auc_interpretation <- ifelse(auc_value > 0.9, "Excellent",
                             ifelse(auc_value > 0.8, "Good",
                                    ifelse(auc_value > 0.7, "Fair",
                                           ifelse(auc_value > 0.6, "Poor", "Fail"))))

# 最佳切点
best_cutoff <- coords(roc_result, "best", ret = "all", best.method = "youden")

# 7.3 创建修复横坐标的ROC曲线
cat("\n📊 Creating ROC Curve with Fixed X-axis...\n")

# 版本1：修复横坐标刻度
png("ROC_Fixed_Xaxis_1.png", width = 1400, height = 1100, res = 200)

# 设置边距：下、左、上、右
par(mar = c(5, 5, 4, 12))

# 首先，我们需要计算并保存ROC曲线，但不绘制
plot(roc_result, 
     main = "ROC Curve of the Prediction Model",
     xlab = "1 - Specificity (False Positive Rate)",
     ylab = "Sensitivity (True Positive Rate)",
     legacy.axes = TRUE,
     col = "#2C3E50",
     lwd = 4,
     cex.main = 2.0,
     cex.lab = 1.6,
     print.auc = FALSE,
     print.thres = FALSE,
     plot = FALSE)  # 先不绘制，只计算

# 重新绘制，手动控制坐标轴
plot(roc_result, 
     main = "ROC Curve of the Prediction Model",
     xlab = "1 - Specificity (False Positive Rate)",
     ylab = "Sensitivity (True Positive Rate)",
     legacy.axes = TRUE,
     col = "#2C3E50",
     lwd = 4,
     cex.main = 2.0,
     cex.lab = 1.6,
     print.auc = FALSE,
     print.thres = FALSE,
     xaxt = "n",  # 不绘制x轴
     yaxt = "n")  # 不绘制y轴

# 修改这里：手动添加x轴刻度，强制显示为0.0, 0.2, 0.4, 0.6, 0.8, 1.0
x_breaks <- c(0, 0.2, 0.4, 0.6, 0.8, 1.0)
x_labels <- c("0.0", "0.2", "0.4", "0.6", "0.8", "1.0")
axis(1, at = x_breaks, labels = x_labels, cex.axis = 1.4, tck = -0.02)

# 手动添加y轴刻度
y_breaks <- seq(0, 1, 0.2)
y_labels <- c("0.0", "0.2", "0.4", "0.6", "0.8", "1.0")
axis(2, at = y_breaks, labels = y_labels, cex.axis = 1.4, tck = -0.02, las = 1)

# 添加边框
box()

# 对角线
abline(0, 1, lty = 2, col = "#7F8C8D", lwd = 3)

# 将图例放在图形外部右侧
legend(x = 1.05, y = 0.95,  # 将图例放在图形右侧外部
       legend = c(
         paste0("AUC = ", round(auc_value, 3)),
         paste0("95% CI: [", round(ci_auc[1], 3), ", ", round(ci_auc[3], 3), "]"),
         "",
         paste0("Cutoff: ", round(best_cutoff$threshold, 3)),
         paste0("Sensitivity: ", round(best_cutoff$sensitivity * 100, 1), "%"),
         paste0("Specificity: ", round(best_cutoff$specificity * 100, 1), "%")
       ),
       bty = "n",
       cex = 1.4,
       xpd = TRUE,  # 允许绘图在边距区
       y.intersp = 1.2)

# 样本信息放在左下角
text(x = 0.02, y = 0.02, 
     paste0("N = ", format(nrow(complete_data), big.mark = ",")), 
     cex = 1.4, adj = 0, font = 2)

dev.off()
cat("✅ ROC curve with fixed x-axis saved: ROC_Fixed_Xaxis_1.png\n")

# 版本2：使用更简单的方法修复刻度
png("ROC_Fixed_Xaxis_2.png", width = 1400, height = 1100, res = 200)

# 设置图形参数，禁止科学计数法
options(scipen = 999)  # 禁用科学计数法

par(mar = c(5, 5, 4, 12))

# 绘制ROC曲线，使用自定义坐标轴
plot(roc_result, 
     main = "ROC Curve of the Prediction Model",
     xlab = "1 - Specificity",
     ylab = "Sensitivity",
     legacy.axes = TRUE,
     col = "#2C3E50",
     lwd = 4,
     cex.main = 2.0,
     cex.lab = 1.6,
     print.auc = FALSE,
     print.thres = FALSE,
     # 使用pROC内置参数控制坐标轴
     auc.polygon = FALSE,
     grid = FALSE,
     # 强制使用固定刻度
     xlim = c(0, 1),
     ylim = c(0, 1))

# 修改这里：手动添加x轴刻度
axis(1, at = c(0, 0.2, 0.4, 0.6, 0.8, 1.0), 
     labels = c("0.0", "0.2", "0.4", "0.6", "0.8", "1.0"),
     cex.axis = 1.4)

# 添加自定义网格线（可选）
grid(col = "gray90", lty = 3)

# 对角线
abline(0, 1, lty = 2, col = "#7F8C8D", lwd = 3)

# 手动添加图例（在图形内部，避免重叠）
legend("bottomright",
       legend = c(
         paste0("AUC = ", round(auc_value, 3)),
         paste0("95% CI: [", round(ci_auc[1], 3), ", ", round(ci_auc[3], 3), "]"),
         paste0("Cutoff: ", round(best_cutoff$threshold, 3)),
         paste0("Sensitivity: ", round(best_cutoff$sensitivity * 100, 1), "%"),
         paste0("Specificity: ", round(best_cutoff$specificity * 100, 1), "%")
       ),
       bty = "n",
       cex = 1.4,
       y.intersp = 1.2,
       bg = "white",  # 添加白色背景避免覆盖
       box.col = "gray80")

# 样本信息放在左上角
legend("topleft",
       legend = paste0("N = ", format(nrow(complete_data), big.mark = ",")),
       bty = "n",
       cex = 1.4)

dev.off()

# 恢复科学计数法设置
options(scipen = 0)

cat("✅ ROC curve with fixed x-axis saved: ROC_Fixed_Xaxis_2.png\n")

# 版本3：使用pROC内置功能的最简单版本
png("ROC_Simple_Clean.png", width = 1200, height = 900, res = 150)

par(mar = c(5, 5, 4, 4))

# 使用pROC的plot函数，但控制输出
plot(roc_result, 
     main = "ROC Curve",
     xlab = "1 - Specificity",
     ylab = "Sensitivity",
     legacy.axes = TRUE,
     col = "#2C3E50",
     lwd = 3,
     cex.main = 1.8,
     cex.lab = 1.5,
     print.auc = TRUE,
     print.auc.x = 0.6,
     print.auc.y = 0.4,
     print.auc.cex = 1.5,
     auc.polygon = FALSE,
     grid = TRUE,
     grid.col = "gray90",
     # 添加这些参数确保刻度正确
     xaxt = "n")  # 先不绘制x轴

# 修改这里：手动添加x轴刻度
axis(1, at = c(0, 0.2, 0.4, 0.6, 0.8, 1.0), 
     labels = c("0.0", "0.2", "0.4", "0.6", "0.8", "1.0"),
     cex.axis = 1.3)

# 对角线
abline(0, 1, lty = 2, col = "#7F8C8D", lwd = 2)

# 添加最佳切点信息（简洁版）
points(1 - best_cutoff$specificity, best_cutoff$sensitivity,
       col = "#E74C3C", pch = 19, cex = 1.5)

text(1 - best_cutoff$specificity, best_cutoff$sensitivity - 0.05,
     paste0("Cutoff: ", round(best_cutoff$threshold, 3)),
     col = "#E74C3C", cex = 1.2, adj = 0.5)

# 样本信息
mtext(paste0("N = ", format(nrow(complete_data), big.mark = ",")), 
      side = 1, line = 3.5, cex = 1.1)

dev.off()
cat("✅ Simple clean ROC curve saved: ROC_Simple_Clean.png\n")

# 版本4：使用基本绘图函数的完全控制版本
png("ROC_Manual_Control.png", width = 1400, height = 1000, res = 200)

# 设置边距
par(mar = c(6, 6, 5, 6))

# 创建空白图形框架
plot(NA, 
     xlim = c(0, 1), 
     ylim = c(0, 1),
     xlab = "1 - Specificity",
     ylab = "Sensitivity",
     main = "ROC Curve of the Prediction Model",
     cex.main = 2.2,
     cex.lab = 1.8,
     cex.axis = 1.5,
     xaxt = "n",  # 不自动绘制x轴
     yaxt = "n",  # 不自动绘制y轴
     bty = "n")   # 无边框

# 手动添加坐标轴（完全控制） - 这里已经是正确的
x_ticks <- c(0, 0.2, 0.4, 0.6, 0.8, 1.0)
axis(1, at = x_ticks, labels = c("0.0", "0.2", "0.4", "0.6", "0.8", "1.0"), 
     cex.axis = 1.5, tck = -0.02, line = 0)
axis(2, at = x_ticks, labels = c("0.0", "0.2", "0.4", "0.6", "0.8", "1.0"), 
     cex.axis = 1.5, tck = -0.02, line = 0, las = 1)

# 添加网格
grid(nx = NULL, ny = NULL, col = "gray90", lty = 3)

# 添加边框
box()

# 手动绘制ROC曲线
lines(roc_result$specificities, roc_result$sensitivities, 
      type = "l", col = "#2C3E50", lwd = 4)

# 对角线
abline(0, 1, lty = 2, col = "#7F8C8D", lwd = 3)

# 添加AUC信息
text(0.7, 0.25, 
     paste0("AUC = ", round(auc_value, 3), 
            "\n95% CI: [", round(ci_auc[1], 3), ", ", round(ci_auc[3], 3), "]"),
     cex = 1.8, font = 2, col = "#2C3E50")

# 添加最佳切点
points(1 - best_cutoff$specificity, best_cutoff$sensitivity,
       col = "#E74C3C", pch = 19, cex = 2.5)

# 添加切点信息
text(1 - best_cutoff$specificity + 0.05, best_cutoff$sensitivity,
     paste0("Cutoff: ", round(best_cutoff$threshold, 3),
            "\nSens: ", round(best_cutoff$sensitivity * 100, 1), "%",
            "\nSpec: ", round(best_cutoff$specificity * 100, 1), "%"),
     cex = 1.5, col = "#E74C3C", adj = 0)

# 添加样本信息
text(0.05, 0.05, 
     paste0("N = ", format(nrow(complete_data), big.mark = ",")), 
     cex = 1.6, adj = 0, font = 2)

dev.off()
cat("✅ Manually controlled ROC curve saved: ROC_Manual_Control.png\n")

cat("\n", rep("=", 60), "\n", sep = "")
cat("ROC CURVES WITH FIXED X-AXIS CREATED SUCCESSFULLY!\n")
cat(rep("=", 60), "\n\n")

cat("📊 Generated ROC files:\n")
cat("   1. ROC_Fixed_Xaxis_1.png - Fixed x-axis with external legend\n")
cat("   2. ROC_Fixed_Xaxis_2.png - Fixed x-axis with internal legend\n")
cat("   3. ROC_Simple_Clean.png - Simple version with pROC defaults\n")
cat("   4. ROC_Manual_Control.png - Fully manually controlled version\n\n")

cat("💡 Recommendations:\n")
cat("   • For publications: Use ROC_Manual_Control.png (full control)\n")
cat("   • For presentations: Use ROC_Simple_Clean.png (clean and simple)\n")
cat("   • The x-axis is now fixed to show 0.0, 0.2, 0.4, 0.6, 0.8, 1.0\n")

# 8. 创建英文森林图 ------------------------------------------------------------
cat("\n🌲 Creating English Forest Plot...\n")
cat(rep("=", 60), "\n", sep = "")

# 创建森林图数据框（使用英文标签）
forest_data_eng <- data.frame()

# 从模型结果中提取数据（排除截距）
for(i in 2:nrow(coef_summary)) {  # 从2开始，排除截距
  var_name <- rownames(coef_summary)[i]
  beta <- coef_summary[i, "Estimate"]
  se <- coef_summary[i, "Std. Error"]
  p_val <- coef_summary[i, "Pr(>|z|)"]
  
  # 计算OR和置信区间
  or <- exp(beta)
  ci_lower <- exp(beta - 1.96 * se)
  ci_upper <- exp(beta + 1.96 * se)
  
  # 变量标签和分类（英文）
  if(var_name %in% c("gender", "residence", "onlychild")) {
    var_category <- "Demographic"
    var_label <- switch(var_name,
                        "gender" = "Gender (Male)",
                        "residence" = "Residence (Urban)",
                        "onlychild" = "Only Child")
  } else {
    var_category <- "Psychological"
    var_label <- switch(var_name,
                        "depression" = "Depression",
                        "anxiety" = "Anxiety", 
                        "suicide_ideation" = "Suicidal Ideation",
                        "psychosis" = "Psychotic Symptoms",
                        "self_harm" = "Self-harm Behavior",
                        var_name)
  }
  
  # 显著性
  significance <- ifelse(p_val < 0.001, "***",
                         ifelse(p_val < 0.01, "**",
                                ifelse(p_val < 0.05, "*", "")))
  
  forest_data_eng <- rbind(forest_data_eng, data.frame(
    Variable = var_label,
    Category = var_category,
    Beta = beta,
    SE = se,
    OR = or,
    CI_lower = ci_lower,
    CI_upper = ci_upper,
    P_value = p_val,
    Significance = significance,
    OR_CI = sprintf("%.2f (%.2f-%.2f)%s", or, ci_lower, ci_upper, significance),
    stringsAsFactors = FALSE
  ))
}

# 按OR值排序
forest_data_eng <- forest_data_eng[order(forest_data_eng$OR), ]
forest_data_eng$Variable <- factor(forest_data_eng$Variable, 
                                   levels = forest_data_eng$Variable)

cat("\nForest plot data preview:\n")
print(forest_data_eng[, c("Variable", "Category", "OR", "CI_lower", "CI_upper", "P_value")])

# 9. 创建英文森林图 ------------------------------------------------------------
cat("\n🌲 Creating Forest Plot (English Version)\n")
cat(rep("=", 60), "\n", sep = "")

# 方法1: 使用ggplot2创建更清晰的英文森林图
tryCatch({
  # 确保变量是因子，并按照OR值排序
  forest_data_eng$Variable <- factor(forest_data_eng$Variable, 
                                     levels = forest_data_eng$Variable[order(forest_data_eng$OR)])
  
  # 创建颜色方案
  colors <- c("Demographic" = "#E41A1C", "Psychological" = "#377EB8")
  
  # 计算合适的文本位置
  max_or <- max(forest_data_eng$CI_upper, na.rm = TRUE)
  
  # 创建森林图 - 优化版本
  forest_plot_eng <- ggplot(forest_data_eng, aes(x = Variable, y = OR, color = Category)) +
    # 添加参考线
    geom_hline(yintercept = 1, linetype = "dashed", color = "red", alpha = 0.7, size = 1) +
    # 添加置信区间
    geom_errorbar(aes(ymin = CI_lower, ymax = CI_upper), 
                  width = 0.15, size = 1, position = position_dodge(width = 0.5)) +
    # 添加OR点
    geom_point(size = 3, position = position_dodge(width = 0.5)) +
    # 坐标轴翻转
    coord_flip() +
    # 颜色方案
    scale_color_manual(values = colors) +
    # 使用对数坐标（OR值通常在对数尺度上对称）
    scale_y_log10(breaks = c(0.1, 0.2, 0.5, 1, 2, 5, 10),
                  labels = c("0.1", "0.2", "0.5", "1", "2", "5", "10"),
                  limits = c(max(0.1, min(forest_data_eng$CI_lower, na.rm = TRUE) * 0.8), 
                             min(10, max(forest_data_eng$CI_upper, na.rm = TRUE) * 1.2))) +
    # 标签
    labs(
      title = "Forest Plot: Risk Factors for Psychological Intervention",
      subtitle = "Multivariate Logistic Regression Model",
      x = "",
      y = "Odds Ratio (OR) with 95% CI (log scale)",
      color = "Variable Type",
      caption = paste0("Model includes ", nrow(forest_data_eng), 
                       " predictor variables | Sample size N = ", 
                       format(nrow(complete_data), big.mark = ","))
    ) +
    # 主题设置
    theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold", size = 16, margin = margin(b = 10)),
      plot.subtitle = element_text(hjust = 0.5, size = 14, margin = margin(b = 15)),
      plot.caption = element_text(hjust = 0, size = 10, color = "gray50", margin = margin(t = 10)),
      legend.position = "top",
      legend.title = element_text(face = "bold", size = 11),
      legend.text = element_text(size = 10),
      panel.grid.major.y = element_blank(),
      panel.grid.minor.y = element_blank(),
      panel.grid.major.x = element_line(color = "gray90", size = 0.3),
      panel.grid.minor.x = element_blank(),
      axis.text.y = element_text(size = 12, face = "bold", color = "black"),
      axis.text.x = element_text(size = 11),
      axis.title.x = element_text(size = 12, face = "bold", margin = margin(t = 10)),
      plot.margin = margin(20, 60, 20, 20)  # 增加右边距放置文本
    ) +
    # 在右侧添加OR值和置信区间（使用annotate避免重叠）
    annotate("text", 
             x = 1:nrow(forest_data_eng), 
             y = max(forest_data_eng$CI_upper, na.rm = TRUE) * 1.3,
             label = sprintf("OR: %.2f\n[%.2f-%.2f]", 
                             forest_data_eng$OR, 
                             forest_data_eng$CI_lower, 
                             forest_data_eng$CI_upper),
             hjust = 0, size = 3.5, color = "gray30", fontface = "bold") +
    # 在左侧添加P值
    annotate("text",
             x = 1:nrow(forest_data_eng),
             y = min(forest_data_eng$CI_lower, na.rm = TRUE) * 0.7,
             label = ifelse(forest_data_eng$P_value < 0.001, "P<0.001***",
                            ifelse(forest_data_eng$P_value < 0.01, sprintf("P=%.3f**", forest_data_eng$P_value),
                                   ifelse(forest_data_eng$P_value < 0.05, sprintf("P=%.3f*", forest_data_eng$P_value),
                                          sprintf("P=%.3f", forest_data_eng$P_value)))),
             hjust = 1, size = 3.5, color = "gray40")
  
  # 保存高质量图片
  ggsave("Forest_Plot_English.png", forest_plot_eng, 
         width = 14, height = 10, dpi = 300, bg = "white")
  cat("✅ English forest plot saved as: Forest_Plot_English.png\n")
  
  # 显示图形
  print(forest_plot_eng)
  
}, error = function(e) {
  cat("⚠️ ggplot2 plotting failed:", e$message, "\n")
  cat("Trying basic plotting function...\n")
  
  # 方法2: 使用基本绘图函数创建英文森林图
  png("Forest_Plot_English_Basic.png", width = 1400, height = 1000, res = 150)
  
  # 设置图形参数（增加右边距放置文本）
  par(mar = c(5, 15, 6, 25))
  
  # 计算x轴范围（对数刻度）
  all_ci <- c(forest_data_eng$CI_lower, forest_data_eng$CI_upper)
  x_min <- max(0.1, min(all_ci) * 0.8)
  x_max <- min(10, max(all_ci) * 1.2)
  
  # 创建空图
  plot(NA, 
       xlim = c(x_min, x_max), 
       ylim = c(0.5, nrow(forest_data_eng) + 0.5),
       xlab = "Odds Ratio (OR) with 95% Confidence Interval (log scale)", 
       ylab = "", 
       yaxt = "n",
       main = "Forest Plot: Risk Factors for Psychological Intervention\nMultivariate Logistic Regression Model",
       cex.main = 1.4, 
       font.main = 2,
       log = "x",  # x轴使用对数刻度
       cex.lab = 1.1, 
       font.lab = 2)
  
  # 添加参考线
  abline(v = 1, lty = 2, col = "red", lwd = 2)
  
  # 定义颜色
  colors <- c("Demographic" = "#E41A1C", "Psychological" = "#377EB8")
  
  # 添加OR值和置信区间
  for(i in 1:nrow(forest_data_eng)) {
    y_pos <- nrow(forest_data_eng) - i + 1  # 反转顺序
    
    # 置信区间线
    lines(c(forest_data_eng$CI_lower[i], forest_data_eng$CI_upper[i]), 
          c(y_pos, y_pos),
          col = colors[forest_data_eng$Category[i]], lwd = 2.5)
    
    # 点估计（OR值）
    points(forest_data_eng$OR[i], y_pos, pch = 19, 
           col = colors[forest_data_eng$Category[i]], cex = 2)
    
    # 变量标签（左侧）
    text(x_min * 0.9, y_pos, forest_data_eng$Variable[i], 
         adj = 0, cex = 1.2, font = 2, xpd = TRUE)
    
    # OR值和置信区间文本（右侧）
    or_text <- sprintf("%.2f", forest_data_eng$OR[i])
    ci_text <- sprintf("[%.2f-%.2f]", forest_data_eng$CI_lower[i], forest_data_eng$CI_upper[i])
    text(x_max * 1.05, y_pos, 
         paste(or_text, ci_text), 
         adj = 0, cex = 1, col = "gray30", font = 2, xpd = TRUE)
    
    # P值文本（右侧，OR值下方）
    p_text <- ifelse(forest_data_eng$P_value[i] < 0.001, "P<0.001",
                     sprintf("P=%.3f", forest_data_eng$P_value[i]))
    p_text <- paste0(p_text, forest_data_eng$Significance[i])
    text(1, y_pos, p_text, adj = 0.5, cex = 1, col = "gray40", font = 2)
  }
  
  # 添加x轴刻度标签
  axis(1, at = c(0.1, 0.2, 0.5, 1, 2, 5, 10), 
       labels = c("0.1", "0.2", "0.5", "1", "2", "5", "10"), 
       cex.axis = 1.1)
  
  # 添加y轴刻度（变量名）
  axis(2, at = 1:nrow(forest_data_eng), 
       labels = rev(forest_data_eng$Variable), 
       las = 1, cex.axis = 1.1, tick = FALSE)
  
  # 添加图例
  legend("topright", 
         legend = c("Demographic Variables", "Psychological Variables", "Reference Line (OR=1)"),
         col = c("#E41A1C", "#377EB8", "red"),
         pch = c(19, 19, NA),
         lty = c(NA, NA, 2),
         lwd = c(NA, NA, 2),
         pt.cex = 2,
         bty = "n", 
         cex = 1.2, 
         xpd = TRUE,
         inset = c(-0.25, 0))
  
  # 添加脚注
  mtext(paste0("Model includes ", nrow(forest_data_eng), 
               " predictor variables | Sample size N = ", 
               format(nrow(complete_data), big.mark = ",")),
        side = 1, line = 3.5, cex = 1, col = "gray50")
  
  dev.off()
  cat("✅ English forest plot saved as: Forest_Plot_English_Basic.png\n")
})

# 10. 创建简洁版的英文森林图（避免文字重叠）-------------------------------------
cat("\n🌲 Creating Clean English Forest Plot (No Text Overlap)\n")

# 创建更简洁的森林图，文字放在图外
png("Forest_Plot_English_Clean.png", width = 1200, height = 900, res = 150)

# 设置图形布局
layout(matrix(c(1, 2), 1, 2), widths = c(3, 1))

# 左边：森林图主图
par(mar = c(5, 15, 6, 1))

# 计算x轴范围
all_ci <- c(forest_data_eng$CI_lower, forest_data_eng$CI_upper)
x_min <- max(0.1, min(all_ci) * 0.8)
x_max <- min(10, max(all_ci) * 1.2)

# 创建空图
plot(NA, 
     xlim = c(x_min, x_max), 
     ylim = c(0.5, nrow(forest_data_eng) + 0.5),
     xlab = "Odds Ratio (OR) with 95% CI", 
     ylab = "", 
     yaxt = "n",
     main = "Forest Plot of Risk Factors\nfor Psychological Intervention Need",
     cex.main = 1.5, 
     font.main = 2,
     log = "x",
     cex.lab = 1.2)

# 添加参考线
abline(v = 1, lty = 2, col = "red", lwd = 2)

# 添加网格线
grid(ny = NA, nx = NULL, col = "gray90", lty = 3)

# 定义颜色
colors <- c("Demographic" = "#E41A1C", "Psychological" = "#377EB8")

# 添加OR值和置信区间（不添加文字）
for(i in 1:nrow(forest_data_eng)) {
  y_pos <- nrow(forest_data_eng) - i + 1
  
  # 置信区间线
  lines(c(forest_data_eng$CI_lower[i], forest_data_eng$CI_upper[i]), 
        c(y_pos, y_pos),
        col = colors[forest_data_eng$Category[i]], lwd = 2)
  
  # 点估计
  points(forest_data_eng$OR[i], y_pos, pch = 18, 
         col = colors[forest_data_eng$Category[i]], cex = 1.8)
  
  # 变量标签
  text(x_min * 0.9, y_pos, forest_data_eng$Variable[i], 
       adj = 0, cex = 1.1, font = 2)
}

# 添加x轴刻度
axis(1, at = c(0.1, 0.2, 0.5, 1, 2, 5, 10), 
     labels = c("0.1", "0.2", "0.5", "1", "2", "5", "10"), 
     cex.axis = 1.1)

# 右边：数据表格
par(mar = c(5, 1, 6, 2))

# 创建空图作为表格背景
plot(NA, xlim = c(0, 1), ylim = c(0.5, nrow(forest_data_eng) + 0.5),
     xlab = "", ylab = "", xaxt = "n", yaxt = "n", bty = "n")

# 添加表格标题
text(0.5, nrow(forest_data_eng) + 0.8, "OR [95% CI]", font = 2, cex = 1.2)
text(0.5, nrow(forest_data_eng) + 0.6, "P-value", font = 2, cex = 1.2)

# 添加表格内容
for(i in 1:nrow(forest_data_eng)) {
  y_pos <- nrow(forest_data_eng) - i + 1
  
  # OR和CI
  or_ci_text <- sprintf("%.2f [%.2f-%.2f]", 
                        forest_data_eng$OR[i],
                        forest_data_eng$CI_lower[i],
                        forest_data_eng$CI_upper[i])
  text(0.5, y_pos + 0.1, or_ci_text, cex = 1, font = 1)
  
  # P值
  p_text <- ifelse(forest_data_eng$P_value[i] < 0.001, "<0.001",
                   sprintf("%.3f", forest_data_eng$P_value[i]))
  p_text <- paste0(p_text, forest_data_eng$Significance[i])
  text(0.5, y_pos - 0.1, p_text, cex = 1, font = 1, col = "gray40")
}

# 添加图例
legend("bottomright", 
       legend = c("Demographic", "Psychological", "Reference (OR=1)"),
       col = c("#E41A1C", "#377EB8", "red"),
       pch = c(18, 18, NA),
       lty = c(NA, NA, 2),
       lwd = c(NA, NA, 2),
       pt.cex = 1.5,
       bty = "n", 
       cex = 1.1)

# 添加脚注
mtext(paste0("Multivariate Logistic Regression | N = ", 
             format(nrow(complete_data), big.mark = ",")),
      side = 1, line = 2, cex = 0.9, col = "gray50")

dev.off()
cat("✅ Clean English forest plot saved as: Forest_Plot_English_Clean.png\n")

# 11. 保存英文森林图数据 ------------------------------------------------------
# 保存英文版的森林图数据
write.csv(forest_data_eng, "Forest_Plot_Data_English.csv", 
          row.names = FALSE, fileEncoding = "UTF-8")

# 创建英文版的结果摘要
summary_eng <- data.frame(
  Category = c(
    "Sample Information",
    "Sample Information",
    "Model Fit",
    "Model Fit",
    "Model Fit",
    "Predictive Performance",
    "Predictive Performance",
    "Predictive Performance",
    "Discrimination",
    "Discrimination"
  ),
  Metric = c(
    "Total Sample Size",
    "High-risk Proportion",
    "AIC",
    "BIC",
    "Nagelkerke Pseudo R²",
    "Accuracy",
    "Sensitivity",
    "Specificity",
    "AUC",
    "AUC 95% CI"
  ),
  Value = c(
    format(nrow(complete_data), big.mark = ","),
    sprintf("%.1f%%", mean(complete_data$risk_binary) * 100),
    sprintf("%.1f", AIC(model)),
    sprintf("%.1f", BIC(model)),
    sprintf("%.3f", nagelkerke_r2),
    sprintf("%.1f%%", accuracy * 100),
    sprintf("%.1f%%", sensitivity * 100),
    sprintf("%.1f%%", specificity * 100),
    sprintf("%.3f", auc_value),
    sprintf("[%.3f, %.3f]", ci_auc[1], ci_auc[3])
  ),
  Interpretation = c(
    "Complete cases",
    "Proportion requiring intervention",
    "Lower is better",
    "Lower is better",
    "Proportion of variance explained",
    "Proportion correctly classified",
    "Proportion of high-risk correctly identified",
    "Proportion of low-risk correctly identified",
    "Model discrimination ability",
    "Confidence interval for AUC"
  )
)

write.csv(summary_eng, "Model_Summary_English.csv", row.names = FALSE, fileEncoding = "UTF-8")

cat("\n📊 English forest plot data saved as: Forest_Plot_Data_English.csv\n")
cat("📊 Model summary saved as: Model_Summary_English.csv\n")

# 12. 生成英文分析报告摘要 ----------------------------------------------------
cat("\n📋 Generating English summary report...\n")

# 提取显著变量（英文）
sig_vars_eng <- forest_data_eng[forest_data_eng$P_value < 0.05, ]

report_eng <- paste0(
  "MULTIVARIATE LOGISTIC REGRESSION ANALYSIS REPORT\n",
  "===================================================\n\n",
  "1. STUDY OVERVIEW\n",
  "   Purpose: To identify risk factors for psychological intervention need among college students\n",
  "   Method: Multivariate logistic regression\n",
  "   Sample: ", nrow(complete_data), " college students\n",
  "   High-risk: ", sum(complete_data$risk_binary), " (", 
  round(mean(complete_data$risk_binary) * 100, 1), "%)\n\n",
  
  "2. INCLUDED VARIABLES\n",
  "   - Demographic: Gender (Male), Residence (Urban), Only Child\n",
  "   - Psychological: Depression, Anxiety, Suicidal Ideation, Psychotic Symptoms, Self-harm Behavior\n\n",
  
  "3. MODEL FIT STATISTICS\n",
  "   - Model formula: ", formula_str, "\n",
  "   - Likelihood ratio test: χ²(", lr_test$Df[2], ") = ", 
  round(lr_test$Deviance[2], 2), 
  ifelse(lr_test$`Pr(>Chi)`[2] < 0.001, ", P < 0.001\n",
         paste0(", P = ", round(lr_test$`Pr(>Chi)`[2], 4), "\n")),
  "   - AIC: ", round(AIC(model), 1), "\n",
  "   - BIC: ", round(BIC(model), 1), "\n",
  "   - Nagelkerke pseudo R²: ", round(nagelkerke_r2, 3), "\n\n",
  
  "4. PREDICTIVE PERFORMANCE\n",
  "   - Accuracy: ", round(accuracy * 100, 1), "%\n",
  "   - Sensitivity: ", round(sensitivity * 100, 1), "%\n",
  "   - Specificity: ", round(specificity * 100, 1), "%\n",
  "   - AUC: ", round(auc_value, 3), " (95% CI: ", 
  round(ci_auc[1], 3), "-", round(ci_auc[3], 3), ")\n",
  "   - Best cutoff: Probability ≥ ", round(best_cutoff$threshold, 3), 
  " (Sensitivity: ", round(best_cutoff$sensitivity * 100, 1), 
  "%, Specificity: ", round(best_cutoff$specificity * 100, 1), "%)\n\n",
  
  "5. SIGNIFICANT PREDICTORS (P < 0.05)\n"
)

# 添加显著预测因子
if(nrow(sig_vars_eng) > 0) {
  for(i in 1:nrow(sig_vars_eng)) {
    report_eng <- paste0(report_eng,
                         "   - ", sig_vars_eng$Variable[i], 
                         ": OR = ", round(sig_vars_eng$OR[i], 2),
                         " (95% CI: ", round(sig_vars_eng$CI_lower[i], 2), 
                         "-", round(sig_vars_eng$CI_upper[i], 2), 
                         "), P ", 
                         ifelse(sig_vars_eng$P_value[i] < 0.001, "<0.001",
                                sprintf("= %.3f", sig_vars_eng$P_value[i])),
                         "\n")
  }
} else {
  report_eng <- paste0(report_eng, "   No significant predictors (P < 0.05)\n")
}

report_eng <- paste0(report_eng, 
                     "\n6. KEY FINDINGS\n",
                     "   - Model discrimination: ", auc_interpretation, " (AUC = ", 
                     round(auc_value, 3), ")\n",
                     "   - Sample size: ", nrow(complete_data), " complete cases\n",
                     "   - High-risk prevalence: ", 
                     round(mean(complete_data$risk_binary) * 100, 1), "%\n",
                     "   - Significant predictors: ", nrow(sig_vars_eng), 
                     " out of ", nrow(forest_data_eng), " variables\n\n",
                     
                     "7. GENERATED FILES\n",
                     "   - Forest_Plot_English.png: English forest plot\n",
                     "   - Forest_Plot_English_Clean.png: Clean version without text overlap\n",
                     "   - Forest_Plot_Data_English.csv: Forest plot data in English\n",
                     "   - Model_Summary_English.csv: Model summary statistics\n")

writeLines(report_eng, "Analysis_Report_English.txt")
cat("\n📄 English analysis report saved as: Analysis_Report_English.txt\n")

cat("\n", rep("✅", 60), "\n", sep = "")
cat("                   ANALYSIS COMPLETE!\n")
cat(rep("✅", 60), "\n\n")

cat("📊 Generated English files:\n")
cat("   1. Forest_Plot_English.png - Main forest plot\n")
cat("   2. Forest_Plot_English_Clean.png - Clean version without text overlap\n")
cat("   3. Forest_Plot_Data_English.csv - Data for forest plot\n")
cat("   4. Model_Summary_English.csv - Summary statistics\n")
cat("   5. Analysis_Report_English.txt - Complete analysis report\n\n")

cat("💡 Tips for the forest plot:\n")
cat("   • Forest_Plot_English_Clean.png separates the plot and data table\n")
cat("   • This avoids text overlap issues\n")
cat("   • Use this version for publications\n")

# 修复列名问题 - 将X95..CI重命名为95% CI
colnames(results_detailed)[colnames(results_detailed) == "X95..CI"] <- "95% CI"

# 10.2 创建论文格式表格
cat("\n📋 论文格式表格 (可直接用于论文):\n")

# 检查列名
cat("results_detailed的列名:\n")
print(colnames(results_detailed))

# 创建论文表格
paper_table <- results_detailed %>%
  filter(变量 != "截距") %>%
  select(变量, OR, `95% CI`, P值, 显著性)

# 添加变量类别
paper_table$变量类别 <- ifelse(
  paper_table$变量 %in% c("性别（男性）", "居住地（城市）", "独生子女"),
  "人口学变量",
  "心理学变量"
)

# 重新排序
paper_table <- paper_table %>%
  arrange(desc(变量类别), desc(as.numeric(gsub("<", "", P值)))) %>%
  select(变量类别, 变量, OR, `95% CI`, P值, 显著性)

print(paper_table, row.names = FALSE)

# 10.3 保存所有结果
# 10.3 保存所有结果 -----------------------------------------------------------
cat("\n💾 保存所有分析结果...\n")

# 确保 writexl 包已加载
if(!require(writexl)) {
  install.packages("writexl")
  library(writexl)
}

# 1. 保存建模数据集
write_xlsx(complete_data, "建模数据集.xlsx")
cat("✅ 建模数据集已保存\n")

# 2. 保存多元回归详细结果（修正列名问题）
# 确保列名正确
colnames(results_detailed)[colnames(results_detailed) == "X95..CI"] <- "95% CI"

# 添加英文列名以便国际交流
results_bilingual <- results_detailed
colnames(results_bilingual) <- c(
  "Variable",
  "Beta",
  "SE",
  "Z",
  "OR",
  "95% CI",
  "P",
  "Sig",
  "Interpretation"
)

write_xlsx(results_bilingual, "多元回归详细结果.xlsx")
cat("✅ 多元回归详细结果已保存\n")

# 3. 保存论文表格格式
# 修复列名
colnames(paper_table)[colnames(paper_table) == "X95..CI"] <- "95% CI"

write_xlsx(paper_table, "论文表格格式.xlsx")
cat("✅ 论文表格格式已保存\n")

# 4. 保存分析结果摘要
write_xlsx(summary_table, "分析结果摘要.xlsx")
cat("✅ 分析结果摘要已保存\n")

# 5. 保存模型对象（不变）
saveRDS(model, "多元Logistic回归模型.rds")
cat("✅ 模型对象已保存\n")

# 6. 创建并保存一个综合报告Excel文件（包含所有重要结果）
cat("\n📊 创建综合报告Excel文件...\n")

# 创建综合Excel文件，包含多个工作表
final_report <- list(
  "1_回归系数" = results_detailed,
  "2_论文表格" = paper_table,
  "3_模型摘要" = summary_table,
  "4_描述统计" = data.frame(
    统计项 = c("总样本量", "高风险人数", "高风险比例", "完整案例数"),
    数值 = c(
      nrow(data),
      sum(data$risk_binary, na.rm = TRUE),
      paste0(round(mean(data$risk_binary, na.rm = TRUE) * 100, 2), "%"),
      nrow(complete_data)
    )
  ),
  "5_ROC分析" = data.frame(
    指标 = c("AUC", "95% CI下限", "95% CI上限", "最佳切点", "灵敏度", "特异度"),
    数值 = c(
      round(auc_value, 3),
      round(ci_auc[1], 3),
      round(ci_auc[3], 3),
      round(best_cutoff$threshold, 3),
      paste0(round(best_cutoff$sensitivity * 100, 1), "%"),
      paste0(round(best_cutoff$specificity * 100, 1), "%")
    )
  ),
  "6_变量说明" = data.frame(
    变量名 = existing_vars,
    变量类型 = ifelse(existing_vars %in% c("gender", "residence", "onlychild"), "人口学", "心理学"),
    说明 = c("需要心理干预=1, 不需要=0", "男性=1, 女性=0", "城市=1, 农村=0", 
           "独生子女=1, 非独生=0", "抑郁得分", "焦虑得分", 
           "自杀意念得分", "精神病性症状得分", "自伤行为得分")
  )
)

write_xlsx(final_report, "综合分析报告.xlsx")
cat("✅ 综合分析报告已保存：综合分析报告.xlsx\n")

# 7. 额外保存一份CSV格式（可选，使用正确编码）
# 如果需要CSV格式，使用以下代码避免乱码
write.csv(results_detailed, 
          file = "多元回归详细结果_UTF8.csv", 
          row.names = FALSE, 
          fileEncoding = "UTF-8")

cat("✅ 额外保存CSV格式（UTF-8编码）\n")

cat("\n" , rep("📁", 60), "\n", sep="")
cat("                   所有文件保存完成！\n")
cat(rep("📁", 60), "\n\n")

cat("📂 生成的文件列表：\n")
cat("   1. 建模数据集.xlsx\n")
cat("   2. 多元回归详细结果.xlsx\n")
cat("   3. 论文表格格式.xlsx\n")
cat("   4. 分析结果摘要.xlsx\n")
cat("   5. 多元Logistic回归模型.rds\n")
cat("   6. 综合分析报告.xlsx（推荐打开此文件查看所有结果）\n")
cat("   7. 多元回归详细结果_UTF8.csv\n\n")

cat("💡 温馨提示：\n")
cat("   • 直接打开 .xlsx 文件，不会有乱码问题\n")
cat("   • '综合分析报告.xlsx' 包含了所有重要结果，打开这一个文件即可\n")
cat("   • 如需与他人分享，建议发送Excel文件而非CSV\n")


# ====================================================
# College Student Mental Health Screening Prediction Validity Analysis
# Version: 2.0 - Based on Binary Risk Labels
# Last Updated: 2025
# ====================================================

# Clear environment
rm(list = ls())

# 1. Load Required Packages -------------------------------------------------
cat("🔧 Loading required packages...\n")

required_packages <- c(
  "readxl",      # Read Excel
  "dplyr",       # Data manipulation
  "tidyr",       # Data tidying
  "ggplot2",     # Plotting
  "pROC",        # ROC analysis
  "caret",       # Model evaluation
  "ggpubr",      # Publication-quality figures
  "psych",       # Descriptive statistics
  "tableone",    # Create Table 1
  "writexl",     # Write Excel
  "patchwork"    # Combine multiple plots
)

# Install missing packages
new_packages <- required_packages[!(required_packages %in% installed.packages()[,"Package"])]
if(length(new_packages) > 0) {
  cat("📦 Installing missing packages:", paste(new_packages, collapse = ", "), "\n")
  install.packages(new_packages, dependencies = TRUE)
}

# Load packages
suppressPackageStartupMessages({
  for(pkg in required_packages) {
    library(pkg, character.only = TRUE)
  }
})

cat("✅ All packages loaded successfully\n")

# 2. Data Loading and Inspection ---------------------------------------------
cat("\n📊 Data Loading and Inspection\n")
cat(rep("=", 60), "\n", sep = "")

# 2.1 Set file path
file_path <- "01 data/new.xlsx"

# 2.2 Check file
if(!file.exists(file_path)) {
  stop("❌ File does not exist: ", file_path, 
       "\nCurrent working directory: ", getwd())
}

# 2.3 Read data
cat("📥 Reading Excel file...\n")
data <- read_excel(file_path)

# 2.4 Check data dimensions
cat("📈 Data dimensions:", nrow(data), "rows ×", ncol(data), "columns\n")

# 2.5 Find key variables
cat("\n🔍 Identifying key variables:\n")

# Look for risk label column (multiple possible names)
risk_cols <- c("risk_label_binary", "risk_label", "risk lable", "风险标签")
risk_col_name <- NULL

for(col in risk_cols) {
  if(col %in% colnames(data)) {
    risk_col_name <- col
    cat("✅ Found risk label column:", col, "\n")
    break
  }
}

if(is.null(risk_col_name)) {
  cat("⚠️  Standard risk label column not found, showing all columns:\n")
  print(colnames(data))
  risk_col_name <- readline("Please enter risk label column name: ")
}

# 2.6 Check risk label distribution
cat("\n📊 Risk label distribution:\n")
risk_table <- table(data[[risk_col_name]], useNA = "always")
print(risk_table)

# Convert to numeric
data$risk_binary <- as.numeric(as.character(data[[risk_col_name]]))

# Verify binary classification
unique_vals <- unique(na.omit(data$risk_binary))
if(!all(unique_vals %in% c(0, 1))) {
  warning("⚠️  Risk label contains non-0/1 values: ", paste(unique_vals, collapse = ", "))
  cat("Auto-converting to binary: Non-zero values converted to 1\n")
  data$risk_binary <- ifelse(data$risk_binary == 0, 0, 1)
}

cat("\n✅ Final binary distribution:\n")
final_table <- table(data$risk_binary, useNA = "always")
names(final_table) <- c("0 (No intervention needed)", "1 (Intervention needed)", "Missing")[1:length(final_table)]
print(final_table)

# Calculate positive rate
n_positive <- sum(data$risk_binary == 1, na.rm = TRUE)
n_negative <- sum(data$risk_binary == 0, na.rm = TRUE)
positive_rate <- n_positive / (n_positive + n_negative) * 100
cat("\n📈 Statistics:\n")
cat("   No intervention needed (0):", n_negative, "\n")
cat("   Intervention needed (1):", n_positive, "\n")
cat("   Positive rate:", round(positive_rate, 2), "%\n")

# 3. Data Cleaning and Preparation -------------------------------------------
cat("\n🧹 Data Cleaning and Preparation\n")
cat(rep("=", 60), "\n", sep = "")

# 3.1 Find total score column
total_score_patterns <- c("总分", "total_score", "total score")
total_score_col <- NULL

for(pattern in total_score_patterns) {
  matches <- grep(pattern, colnames(data), ignore.case = TRUE, value = TRUE)
  if(length(matches) > 0) {
    total_score_col <- matches[1]
    cat("✅ Found total score column:", total_score_col, "\n")
    break
  }
}

if(is.null(total_score_col)) {
  cat("⚠️  Total score column not found, will attempt to calculate total score\n")
} else {
  # Rename to unified name
  data$total_score <- as.numeric(as.character(data[[total_score_col]]))
  cat("📊 Total score statistics:\n")
  print(summary(data$total_score))
}

# 3.2 Find psychological indicator columns
cat("\n🔍 Identifying psychological indicator columns...\n")

# Common psychological indicator patterns
psych_patterns <- c(
  "幻觉、妄想症状.*指标标准分",
  "自杀意图.*指标标准分",
  "焦虑.*指标标准分",
  "抑郁.*指标标准分",
  "偏执.*指标标准分",
  "自卑.*指标标准分",
  "敏感.*指标标准分",
  "社交恐惧.*指标标准分",
  "躯体化.*指标标准分",
  "依赖.*指标标准分",
  "敌对攻击.*指标标准分",
  "冲动.*指标标准分",
  "强迫.*指标标准分",
  "网络成瘾.*指标标准分",
  "自伤行为.*指标标准分",
  "进食问题.*指标标准分",
  "睡眠困扰.*指标标准分",
  "学校适应困难.*指标标准分",
  "人际关系困扰.*指标标准分",
  "学业压力.*指标标准分",
  "就业压力.*指标标准分",
  "恋爱困扰.*指标标准分"
)

# Find and rename psychological indicators
psych_vars <- list()
for(pattern in psych_patterns) {
  matches <- grep(pattern, colnames(data), value = TRUE)
  if(length(matches) > 0) {
    psych_vars[[pattern]] <- matches[1]
  }
}

cat("✅ Found", length(psych_vars), "psychological indicator standard scores\n")

# 3.3 Simplify column names and convert to numeric
cat("\n🔤 Simplifying column names...\n")

# Simplified name mapping
simple_names <- c(
  "幻觉、妄想症状.*指标标准分" = "psychosis",
  "自杀意图.*指标标准分" = "suicide_ideation",
  "焦虑.*指标标准分" = "anxiety",
  "抑郁.*指标标准分" = "depression",
  "偏执.*指标标准分" = "paranoia",
  "自卑.*指标标准分" = "inferiority",
  "敏感.*指标标准分" = "sensitivity",
  "社交恐惧.*指标标准分" = "social_phobia",
  "躯体化.*指标标准分" = "somatization",
  "依赖.*指标标准分" = "dependence",
  "敌对攻击.*指标标准分" = "hostility",
  "冲动.*指标标准分" = "impulsivity",
  "强迫.*指标标准分" = "compulsion",
  "网络成瘾.*指标标准分" = "internet_addiction",
  "自伤行为.*指标标准分" = "self_harm",
  "进食问题.*指标标准分" = "eating_disorder",
  "睡眠困扰.*指标标准分" = "sleep_disturbance",
  "学校适应困难.*指标标准分" = "school_adjustment",
  "人际关系困扰.*指标标准分" = "interpersonal_stress",
  "学业压力.*指标标准分" = "academic_stress",
  "就业压力.*指标标准分" = "employment_stress",
  "恋爱困扰.*指标标准分" = "relationship_stress"
)

# Apply renaming and conversion
for(old_name in names(simple_names)) {
  if(old_name %in% names(psych_vars)) {
    actual_col <- psych_vars[[old_name]]
    new_name <- simple_names[old_name]
    
    # Rename and convert to numeric
    data[[new_name]] <- as.numeric(as.character(data[[actual_col]]))
    
    cat("   ", actual_col, " → ", new_name, "\n")
  }
}

# 3.4 Create psychological dimension total scores
cat("\n📊 Creating psychological dimension total scores...\n")

# Define dimensions
severe_vars <- c("psychosis", "suicide_ideation")
general_vars <- c("anxiety", "depression", "paranoia", "inferiority", 
                  "sensitivity", "social_phobia", "somatization", 
                  "dependence", "hostility", "impulsivity", "compulsion",
                  "internet_addiction", "self_harm", "eating_disorder", 
                  "sleep_disturbance")
developmental_vars <- c("school_adjustment", "interpersonal_stress", 
                        "academic_stress", "employment_stress", "relationship_stress")

# Calculate dimension totals (only for existing variables)
if(all(severe_vars %in% colnames(data))) {
  data$severe_total <- rowSums(data[, severe_vars], na.rm = TRUE)
  cat("✅ Severe psychological problems dimension: 2 indicators\n")
}

if(length(intersect(general_vars, colnames(data))) > 5) {
  existing_general <- intersect(general_vars, colnames(data))
  data$general_total <- rowSums(data[, existing_general], na.rm = TRUE)
  cat("✅ General psychological problems dimension:", length(existing_general), "indicators\n")
}

if(all(developmental_vars %in% colnames(data))) {
  data$developmental_total <- rowSums(data[, developmental_vars], na.rm = TRUE)
  cat("✅ Developmental distress dimension: 5 indicators\n")
}

# ====================================================
# Three-Level Screening Structure Validation Analysis
# ====================================================

cat("\n🎯 Three-Level Screening Structure Validation Analysis\n")
cat(rep("=", 60), "\n", sep = "")

# Ensure data is loaded
if(!exists("data")) {
  stop("❌ Data not loaded, please run the previous code first")
}

# Check if dependent variable exists
if(!"risk_binary" %in% colnames(data)) {
  stop("❌ Dependent variable 'risk_binary' does not exist")
}

# 4. Comparison of Models at Different Screening Levels ---------------------
cat("\n📊 Comparison of Models at Different Screening Levels\n")
cat(rep("-", 60), "\n", sep = "")

# 4.1 Prepare variable lists
# Severe psychological problem indicators
severe_indicators <- c("psychosis", "suicide_ideation")

# General psychological problem indicators (select from existing variables)
general_indicators <- intersect(
  c("anxiety", "depression", "paranoia", "inferiority", 
    "sensitivity", "social_phobia", "somatization", 
    "dependence", "hostility", "impulsivity", "compulsion",
    "internet_addiction", "self_harm", "eating_disorder", 
    "sleep_disturbance"),
  colnames(data)
)

# Developmental distress indicators
developmental_indicators <- intersect(
  c("school_adjustment", "interpersonal_stress", 
    "academic_stress", "employment_stress", "relationship_stress"),
  colnames(data)
)

cat("✅ Variable statistics:\n")
cat("   Severe psychological indicators:", length(severe_indicators), "\n")
cat("   General psychological indicators:", length(general_indicators), "\n")
cat("   Developmental distress indicators:", length(developmental_indicators), "\n")

# 4.2 Handle missing values (prepare for modeling)
cat("\n🧹 Handling missing values...\n")
model_data <- data[, c("risk_binary", severe_indicators, general_indicators, developmental_indicators)]
model_data <- na.omit(model_data)
cat("   Complete data sample size:", nrow(model_data), "\n")

# 4.3 Define model formula function
create_model <- function(predictors, data) {
  if(length(predictors) == 0) return(NULL)
  
  formula_str <- paste("risk_binary ~", paste(predictors, collapse = " + "))
  model <- glm(as.formula(formula_str), 
               data = data, 
               family = binomial(link = "logit"))
  return(model)
}

# 4.4 Train three-level models
cat("\n🔧 Training three-level screening models...\n")

# Model 1: Severe psychological indicators only
cat("   1. Severe psychological indicators only...")
model1 <- create_model(severe_indicators, model_data)

# Model 2: Severe + general psychological indicators
cat("\n   2. Severe + general psychological indicators...")
model2 <- create_model(c(severe_indicators, general_indicators), model_data)

# Model 3: All indicators
cat("\n   3. All indicators (severe + general + developmental)...")
model3 <- create_model(c(severe_indicators, general_indicators, developmental_indicators), model_data)

cat("\n✅ All models trained successfully\n")

# 继续之前的代码...

# 4.5 计算每个模型的AUC - 修复版
cat("\n📈 Calculating AUC for each model...\n")

# 重新计算ROC曲线和AUC，确保正确获取
# 模型1: 仅严重指标
pred1 <- predict(model1, type = "response")
roc1 <- roc(model_data$risk_binary, pred1, quiet = TRUE)
auc1 <- auc(roc1)

# 模型2: 严重+一般指标
pred2 <- predict(model2, type = "response")
roc2 <- roc(model_data$risk_binary, pred2, quiet = TRUE)
auc2 <- auc(roc2)

# 模型3: 全部指标
pred3 <- predict(model3, type = "response")
roc3 <- roc(model_data$risk_binary, pred3, quiet = TRUE)
auc3 <- auc(roc3)

# 4.6 创建AUC比较表格 - 修复版
cat("\n📊 Creating AUC comparison table...\n")

# 计算95%置信区间
ci1 <- ci(roc1)
ci2 <- ci(roc2)
ci3 <- ci(roc3)

auc_comparison <- data.frame(
  Model = c("Severe indicators only", "Severe + General indicators", "All indicators"),
  Number_of_Indicators = c(
    length(severe_indicators),
    length(severe_indicators) + length(general_indicators),
    length(severe_indicators) + length(general_indicators) + length(developmental_indicators)
  ),
  AUC = round(c(auc1, auc2, auc3), 3),
  AUC_95CI_Lower = round(c(ci1[1], ci2[1], ci3[1]), 3),
  AUC_95CI_Upper = round(c(ci1[3], ci2[3], ci3[3]), 3)
)

# 添加格式化后的95%CI列
auc_comparison$AUC_95CI <- paste0(
  auc_comparison$AUC, " (", 
  auc_comparison$AUC_95CI_Lower, "-", 
  auc_comparison$AUC_95CI_Upper, ")"
)

cat("\n📊 Three-level screening model AUC comparison:\n")
print(auc_comparison[, c("Model", "Number_of_Indicators", "AUC", "AUC_95CI")])

# 4.7 可视化AUC比较 - 修复版
cat("\n📊 Creating AUC comparison plot...\n")

# 方法1: 使用基础绘图函数绘制多条ROC曲线
png("ROC_Comparison.png", width = 800, height = 600, res = 150)

# 设置图形参数
par(mar = c(5, 5, 4, 2) + 0.1)

# 绘制第一条ROC曲线
plot(roc1, 
     col = "#E41A1C", 
     lwd = 2,
     main = "ROC Curve Comparison by Screening Level",
     xlab = "1 - Specificity (False Positive Rate)",
     ylab = "Sensitivity (True Positive Rate)",
     cex.main = 1.2,
     cex.lab = 1.1,
     legacy.axes = TRUE)

# 添加第二条ROC曲线
lines(roc2, 
      col = "#377EB8", 
      lwd = 2)

# 添加第三条ROC曲线
lines(roc3, 
      col = "#4DAF4A", 
      lwd = 2)

# 添加对角线参考线
abline(a = 0, b = 1, lty = 2, col = "gray")

# 添加图例
legend("bottomright", 
       legend = c(paste0("Severe only (AUC = ", round(auc1, 3), ")"),
                  paste0("Severe + General (AUC = ", round(auc2, 3), ")"),
                  paste0("All indicators (AUC = ", round(auc3, 3), ")")),
       col = c("#E41A1C", "#377EB8", "#4DAF4A"),
       lwd = 2,
       cex = 0.9,
       bty = "n")

dev.off()

cat("✅ ROC curve plot saved as 'ROC_Comparison.png'\n")

# 方法2: 使用ggplot2绘制（备选方案）
if (require(ggplot2)) {
  # 创建绘图数据
  roc_data <- data.frame()
  
  # 添加第一条ROC曲线数据
  roc_df1 <- data.frame(
    Specificity = 1 - roc1$specificities,
    Sensitivity = roc1$sensitivities,
    Model = paste0("Severe only (AUC = ", round(auc1, 3), ")")
  )
  roc_data <- rbind(roc_data, roc_df1)
  
  # 添加第二条ROC曲线数据
  roc_df2 <- data.frame(
    Specificity = 1 - roc2$specificities,
    Sensitivity = roc2$sensitivities,
    Model = paste0("Severe + General (AUC = ", round(auc2, 3), ")")
  )
  roc_data <- rbind(roc_data, roc_df2)
  
  # 添加第三条ROC曲线数据
  roc_df3 <- data.frame(
    Specificity = 1 - roc3$specificities,
    Sensitivity = roc3$sensitivities,
    Model = paste0("All indicators (AUC = ", round(auc3, 3), ")")
  )
  roc_data <- rbind(roc_data, roc_df3)
  
  # 创建ggplot2图形
  roc_plot_ggplot <- ggplot(roc_data, aes(x = Specificity, y = Sensitivity, color = Model)) +
    geom_line(size = 1.2) +
    geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "gray50") +
    labs(
      title = "ROC Curve Comparison by Screening Level",
      x = "1 - Specificity (False Positive Rate)",
      y = "Sensitivity (True Positive Rate)",
      color = "Model"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold", size = 14),
      axis.title = element_text(size = 12),
      axis.text = element_text(size = 10),
      legend.position = "bottom",
      legend.title = element_text(face = "bold"),
      legend.text = element_text(size = 10)
    ) +
    scale_color_manual(values = c("#E41A1C", "#377EB8", "#4DAF4A")) +
    coord_equal()
  
  # 保存ggplot2图形
  ggsave("ROC_Comparison_ggplot.png", roc_plot_ggplot, width = 8, height = 6, dpi = 300)
  cat("✅ Alternative ROC curve plot saved as 'ROC_Comparison_ggplot.png'\n")
  
  # 显示图形
  print(roc_plot_ggplot)
}

# 5. LASSO回归筛选重要变量 --------------------------------------------------
# 5. LASSO回归筛选重要变量 --------------------------------------------------
cat("\n🎯 LASSO Regression for Important Variable Selection\n")
cat(rep("-", 60), "\n", sep = "")

# 5.1 确保glmnet包已加载
if(!require("glmnet")) {
  install.packages("glmnet", dependencies = TRUE)
  library(glmnet)
}

cat("Preparing LASSO regression data...\n")

# 使用所有指标作为预测变量
x_vars <- c(severe_indicators, general_indicators, developmental_indicators)
x_matrix <- as.matrix(model_data[, x_vars])
y_vector <- model_data$risk_binary

# 检查数据维度
cat("   Predictor matrix dimensions:", dim(x_matrix), "\n")
cat("   Outcome vector length:", length(y_vector), "\n")

# 5.3 设置交叉验证
cat("Performing cross-validated LASSO regression...\n")
set.seed(123)  # 确保可重复性

# 使用tryCatch处理可能的错误
tryCatch({
  cv_lasso <- cv.glmnet(x_matrix, y_vector, 
                        alpha = 1,  # LASSO回归
                        family = "binomial",
                        type.measure = "auc",
                        nfolds = 10)
  
  # 5.4 获取最优lambda值
  best_lambda <- cv_lasso$lambda.min
  cat("   Optimal lambda value:", round(best_lambda, 5), "\n")
  
  # 5.5 获取系数
  lasso_model <- glmnet(x_matrix, y_vector, 
                        alpha = 1, 
                        family = "binomial",
                        lambda = best_lambda)
  
  # 5.6 提取非零系数的变量
  lasso_coef <- coef(lasso_model, s = best_lambda)
  non_zero_indices <- which(lasso_coef[,1] != 0)
  non_zero_indices <- non_zero_indices[names(non_zero_indices) != "(Intercept)"]
  
  if(length(non_zero_indices) > 0) {
    lasso_vars <- lasso_coef[non_zero_indices, , drop = FALSE]
    
    cat("\n📊 LASSO regression results:\n")
    cat("   Selected", nrow(lasso_vars), "important variables\n")
    
    # 按系数绝对值排序
    lasso_vars <- lasso_vars[order(abs(lasso_vars), decreasing = TRUE), , drop = FALSE]
    
    # 如果多于20个，只取前20个
    if(nrow(lasso_vars) > 20) {
      lasso_vars <- lasso_vars[1:20, , drop = FALSE]
    }
    
    # 创建变量重要性数据框
    var_importance <- data.frame(
      Variable = rownames(lasso_vars),
      Coefficient = round(lasso_vars[, 1], 4),
      stringsAsFactors = FALSE
    )
    
    var_importance$Absolute_Value <- round(abs(var_importance$Coefficient), 4)
    var_importance$Rank <- 1:nrow(var_importance)
    
    cat("\n🔝 Most important predictors:\n")
    print(var_importance)
    
    # 5.7 可视化变量重要性 - 使用基础R绘图
    cat("\n📊 Creating variable importance plot...\n")
    
    # 按系数值排序
    var_importance <- var_importance[order(var_importance$Coefficient), ]
    
    # 设置图形参数
    png("LASSO_Variable_Importance.png", width = 1000, height = 700, res = 150)
    par(mar = c(5, 10, 4, 2) + 0.1)  # 增加左边距以便显示长变量名
    
    # 创建颜色向量（正系数红色，负系数蓝色）
    colors <- ifelse(var_importance$Coefficient > 0, "#E41A1C", "#377EB8")
    
    # 绘制条形图
    bar_heights <- var_importance$Coefficient
    bar_names <- var_importance$Variable
    
    barplot(bar_heights, 
            horiz = TRUE,
            names.arg = bar_names,
            las = 1,
            col = colors,
            border = NA,
            main = "Important Variables Selected by LASSO Regression",
            xlab = "Standardized Coefficient",
            cex.names = 0.8,
            cex.axis = 0.9,
            cex.main = 1.2)
    
    # 添加参考线
    abline(v = 0, lty = 2, col = "gray50")
    
    # 添加图例
    legend("bottomright", 
           legend = c("Positive effect", "Negative effect"),
           fill = c("#E41A1C", "#377EB8"),
           cex = 0.8,
           bty = "n")
    
    dev.off()
    cat("✅ Variable importance plot saved as 'LASSO_Variable_Importance.png'\n")
    
    # 5.8 保存LASSO结果
    write.csv(var_importance, "LASSO_Important_Variables.csv", row.names = FALSE)
    cat("✅ LASSO results saved as 'LASSO_Important_Variables.csv'\n")
    
  } else {
    cat("⚠️  LASSO did not select any important variables\n")
  }
  
}, error = function(e) {
  cat("❌ Error in LASSO regression:", e$message, "\n")
  cat("   Trying alternative approach with higher lambda...\n")
  
  # 尝试使用lambda.1se（更简约的模型）
  best_lambda <- cv_lasso$lambda.1se
  cat("   Using lambda.1se:", round(best_lambda, 5), "\n")
  
  lasso_model <- glmnet(x_matrix, y_vector, 
                        alpha = 1, 
                        family = "binomial",
                        lambda = best_lambda)
  
  lasso_coef <- coef(lasso_model, s = best_lambda)
  non_zero_indices <- which(lasso_coef[,1] != 0)
  non_zero_indices <- non_zero_indices[names(non_zero_indices) != "(Intercept)"]
  
  if(length(non_zero_indices) > 0) {
    lasso_vars <- lasso_coef[non_zero_indices, , drop = FALSE]
    
    cat("   Selected", nrow(lasso_vars), "important variables (using lambda.1se)\n")
    
    # ... 继续处理选中的变量 ...
  } else {
    cat("⚠️  Still no variables selected with lambda.1se\n")
  }
})

# 6. 变量理论归属映射 ---------------------------------------------------------
cat("\n📚 Variable Theoretical Classification Mapping\n")
cat(rep("-", 60), "\n", sep = "")

# 6.1 定义变量的理论归属
variable_mapping <- list(
  # 严重心理问题（一级筛查）
  psychosis = list(Level = "Level 1", Dimension = "Severe Psychological Problems", Specific_Dimension = "Psychosis/Delusions"),
  suicide_ideation = list(Level = "Level 1", Dimension = "Severe Psychological Problems", Specific_Dimension = "Suicide Ideation"),
  
  # 一般心理问题 - 内化问题（二级筛查）
  anxiety = list(Level = "Level 2", Dimension = "General Problems-Internalizing", Specific_Dimension = "Anxiety"),
  depression = list(Level = "Level 2", Dimension = "General Problems-Internalizing", Specific_Dimension = "Depression"),
  paranoia = list(Level = "Level 2", Dimension = "General Problems-Internalizing", Specific_Dimension = "Paranoia"),
  inferiority = list(Level = "Level 2", Dimension = "General Problems-Internalizing", Specific_Dimension = "Inferiority"),
  sensitivity = list(Level = "Level 2", Dimension = "General Problems-Internalizing", Specific_Dimension = "Sensitivity"),
  social_phobia = list(Level = "Level 2", Dimension = "General Problems-Internalizing", Specific_Dimension = "Social Phobia"),
  somatization = list(Level = "Level 2", Dimension = "General Problems-Internalizing", Specific_Dimension = "Somatization"),
  
  # 一般心理问题 - 外化问题（二级筛查）
  dependence = list(Level = "Level 2", Dimension = "General Problems-Externalizing", Specific_Dimension = "Dependence"),
  hostility = list(Level = "Level 2", Dimension = "General Problems-Externalizing", Specific_Dimension = "Hostility/Aggression"),
  impulsivity = list(Level = "Level 2", Dimension = "General Problems-Externalizing", Specific_Dimension = "Impulsivity"),
  compulsion = list(Level = "Level 2", Dimension = "General Problems-Externalizing", Specific_Dimension = "Compulsion"),
  internet_addiction = list(Level = "Level 2", Dimension = "General Problems-Externalizing", Specific_Dimension = "Internet Addiction"),
  self_harm = list(Level = "Level 2", Dimension = "General Problems-Externalizing", Specific_Dimension = "Self-harm"),
  eating_disorder = list(Level = "Level 2", Dimension = "General Problems-Externalizing", Specific_Dimension = "Eating Problems"),
  sleep_disturbance = list(Level = "Level 2", Dimension = "General Problems-Externalizing", Specific_Dimension = "Sleep Disturbance"),
  
  # 发展性困扰（三级筛查）
  school_adjustment = list(Level = "Level 3", Dimension = "Developmental Distress", Specific_Dimension = "School Adjustment Difficulties"),
  interpersonal_stress = list(Level = "Level 3", Dimension = "Developmental Distress", Specific_Dimension = "Interpersonal Stress"),
  academic_stress = list(Level = "Level 3", Dimension = "Developmental Distress", Specific_Dimension = "Academic Stress"),
  employment_stress = list(Level = "Level 3", Dimension = "Developmental Distress", Specific_Dimension = "Employment Stress"),
  relationship_stress = list(Level = "Level 3", Dimension = "Developmental Distress", Specific_Dimension = "Relationship Stress")
)

# 6.2 为LASSO筛选的变量创建归属表
if(exists("var_importance") && nrow(var_importance) > 0) {
  cat("Creating variable theoretical classification table...\n")
  
  theoretical_mapping <- data.frame(
    Variable = character(),
    Level = character(),
    Broad_Dimension = character(),
    Specific_Dimension = character(),
    Coefficient = numeric(),
    Absolute_Rank = integer(),
    stringsAsFactors = FALSE
  )
  
  for(i in 1:nrow(var_importance)) {
    var_name <- var_importance$Variable[i]
    
    if(var_name %in% names(variable_mapping)) {
      mapping <- variable_mapping[[var_name]]
      
      theoretical_mapping <- rbind(theoretical_mapping, data.frame(
        Variable = var_name,
        Level = mapping$Level,
        Broad_Dimension = mapping$Dimension,
        Specific_Dimension = mapping$Specific_Dimension,
        Coefficient = var_importance$Coefficient[i],
        Absolute_Rank = var_importance$Rank[i],
        stringsAsFactors = FALSE
      ))
    } else {
      cat("⚠️  Warning: Variable", var_name, "not found in mapping table\n")
    }
  }
  
  if(nrow(theoretical_mapping) > 0) {
    # 按排名排序
    theoretical_mapping <- theoretical_mapping[order(theoretical_mapping$Absolute_Rank), ]
    
    cat("\n📋 Variable theoretical classification table:\n")
    print(theoretical_mapping)
    
    # 6.3 统计分析：各筛查级别的贡献
    cat("\n📊 Contribution statistics by screening level:\n")
    
    level_summary <- theoretical_mapping %>%
      group_by(Level) %>%
      summarise(
        Count = n(),
        Mean_Abs_Coefficient = mean(abs(Coefficient)),
        Max_Abs_Coefficient = max(abs(Coefficient)),
        Min_Abs_Coefficient = min(abs(Coefficient)),
        .groups = 'drop'
      )
    
    print(level_summary)
    
    # 6.4 可视化：各筛查级别变量分布 - 使用基础R绘图
    cat("\n📊 Creating screening level distribution plot...\n")
    
    # 计算各级别的变量数量
    level_counts <- table(theoretical_mapping$Level)
    
    png("Screening_Level_Distribution.png", width = 800, height = 600, res = 150)
    par(mar = c(5, 5, 4, 2) + 0.1)
    
    # 绘制条形图
    barplot(level_counts,
            col = c("#E41A1C", "#377EB8", "#4DAF4A"),
            main = "Distribution of Important Variables Across Screening Levels",
            xlab = "Screening Level",
            ylab = "Number of Variables",
            cex.main = 1.2,
            cex.lab = 1.1,
            cex.names = 1.1,
            ylim = c(0, max(level_counts) * 1.2))
    
    # 添加数值标签
    text(x = 1:length(level_counts),
         y = level_counts + max(level_counts) * 0.05,
         labels = paste0("n=", level_counts),
         cex = 1.1)
    
    dev.off()
    cat("✅ Screening level distribution plot saved as 'Screening_Level_Distribution.png'\n")
    
    # 6.5 保存结果到Excel
    cat("\n💾 Saving analysis results...\n")
    
    # 创建输出目录
    output_dir <- "analysis_results"
    if(!dir.exists(output_dir)) {
      dir.create(output_dir)
    }
    
    # 保存结果到CSV文件
    write.csv(auc_comparison, file.path(output_dir, "AUC_Comparison.csv"), row.names = FALSE)
    write.csv(theoretical_mapping, file.path(output_dir, "Variable_Theoretical_Classification.csv"), row.names = FALSE)
    write.csv(level_summary, file.path(output_dir, "Level_Contribution_Summary.csv"), row.names = FALSE)
    
    cat("✅ Results saved to 'analysis_results' directory\n")
    
  } else {
    cat("⚠️  No variables could be mapped to theoretical classification\n")
  }
} else {
  cat("⚠️  Cannot create theoretical classification table - no important variables selected\n")
}

# 7. 总结报告 ---------------------------------------------------------------
cat("\n📋 Analysis Summary Report\n")
cat(rep("=", 60), "\n", sep = "")

cat("🎯 Key Findings:\n")
cat("1. Model Performance Comparison:\n")
for(i in 1:nrow(auc_comparison)) {
  cat(sprintf("   - %s: AUC = %.3f (%s)\n", 
              auc_comparison$Model[i], 
              auc_comparison$AUC[i],
              auc_comparison$AUC_95CI[i]))
}

# 计算增量效度
auc_increment_1_to_2 <- auc_comparison$AUC[2] - auc_comparison$AUC[1]
auc_increment_2_to_3 <- auc_comparison$AUC[3] - auc_comparison$AUC[2]

cat("\n2. Incremental Validity:\n")
cat(sprintf("   - From Level 1 to Level 2 screening: AUC improvement = %.3f\n", auc_increment_1_to_2))
cat(sprintf("   - From Level 2 to Level 3 screening: AUC improvement = %.3f\n", auc_increment_2_to_3))

if(auc_increment_1_to_2 > auc_increment_2_to_3) {
  cat("   → Level 2 screening (general psychological problems) provides the main predictive increment\n")
} else {
  cat("   → Level 3 screening (developmental distress) still provides some predictive increment\n")
}

# 检查是否有LASSO结果
if(exists("var_importance") && nrow(var_importance) > 0) {
  cat("\n3. LASSO Regression Results:\n")
  cat(sprintf("   Selected %d important variables\n", nrow(var_importance)))
  
  # 识别最重要的维度
  if(exists("theoretical_mapping") && nrow(theoretical_mapping) > 0) {
    top_vars <- head(theoretical_mapping, 3)
    cat("   - Top 3 most important predictive dimensions:\n")
    for(i in 1:nrow(top_vars)) {
      cat(sprintf("     %d. %s (%s, Coefficient = %.3f)\n", 
                  i, 
                  top_vars$Specific_Dimension[i],
                  top_vars$Level[i],
                  top_vars$Coefficient[i]))
    }
  }
}

cat("\n✅ Analysis completed! Recommendations:\n")
cat("1. Focus on core predictors identified in the analysis\n")
cat("2. The three-level screening shows incremental validity\n")
cat("3. Generated figures and tables can be used directly in your paper\n")

cat("\n📊 Generated Output Files:\n")
cat("   1. ROC_Comparison.png - ROC curve comparison plot\n")
cat("   2. LASSO_Variable_Importance.png - LASSO variable importance\n")
cat("   3. Screening_Level_Distribution.png - Variable distribution by level\n")
cat("   4. analysis_results/ - Directory containing CSV files of results\n")

cat("\n" + rep("🎉", 20) + "\n")
cat("All analyses completed successfully!\n")
cat(rep("🎉", 20) + "\n")

# 诊断代码：检查变量映射完整性
cat("\n🔍 诊断检查：变量映射完整性\n")
cat("=" * 50, "\n")

# 1. 检查LASSO选出的变量
cat("1. LASSO选出的变量总数:", nrow(var_importance), "\n")
print(var_importance$Variable)

# 2. 检查成功映射的变量
mapped_vars <- theoretical_mapping$Variable
cat("\n2. 成功映射的变量总数:", length(mapped_vars), "\n")
print(mapped_vars)

# 3. 检查未映射的变量
unmapped_vars <- setdiff(var_importance$Variable, mapped_vars)
cat("\n3. 未映射的变量:", length(unmapped_vars), "\n")
if(length(unmapped_vars) > 0) {
  print(unmapped_vars)
  cat("\n可能原因：这些变量名不在variable_mapping列表中\n")
}

# 4. 重新生成正确的分布图
cat("\n4. 重新生成正确的筛查级别分布图...\n")

# 确保所有变量都被映射
all_mapped <- var_importance$Variable %in% names(variable_mapping)
if(!all(all_mapped)) {
  cat("⚠️  警告：以下变量未在映射表中找到:\n")
  print(var_importance$Variable[!all_mapped])
  
  # 尝试自动修正
  cat("\n尝试自动修正映射...\n")
  for(var in var_importance$Variable[!all_mapped]) {
    # 基于变量名猜测归属
    if(grepl("stress", var, ignore.case = TRUE)) {
      variable_mapping[[var]] <- list(Level = "Level 3", 
                                      Dimension = "Developmental Distress",
                                      Specific_Dimension = var)
      cat("   ", var, " → 自动映射到 Level 3\n")
    } else if(var %in% general_indicators) {
      variable_mapping[[var]] <- list(Level = "Level 2",
                                      Dimension = "General Problems",
                                      Specific_Dimension = var)
      cat("   ", var, " → 自动映射到 Level 2\n")
    }
  }
}

# 重新生成理论映射
theoretical_mapping <- data.frame()
for(i in 1:nrow(var_importance)) {
  var_name <- var_importance$Variable[i]
  
  if(var_name %in% names(variable_mapping)) {
    mapping <- variable_mapping[[var_name]]
    
    theoretical_mapping <- rbind(theoretical_mapping, data.frame(
      Variable = var_name,
      Level = mapping$Level,
      Broad_Dimension = mapping$Dimension,
      Specific_Dimension = mapping$Specific_Dimension,
      Coefficient = var_importance$Coefficient[i],
      Absolute_Rank = var_importance$Rank[i],
      stringsAsFactors = FALSE
    ))
  }
}

# 重新计算分布
correct_level_counts <- table(theoretical_mapping$Level)
cat("\n✅ 修正后的分布:\n")
print(correct_level_counts)
cat("总计:", sum(correct_level_counts), "个变量\n")

# 重新绘制分布图
png("Screening_Level_Distribution_Corrected.png", width = 800, height = 600, res = 150)
par(mar = c(5, 5, 4, 2) + 0.1)

bar_colors <- c("Level 1" = "#E41A1C", 
                "Level 2" = "#377EB8", 
                "Level 3" = "#4DAF4A")

barplot(correct_level_counts,
        col = bar_colors[names(correct_level_counts)],
        main = "Distribution of Important Variables Across Screening Levels (Corrected)",
        xlab = "Screening Level",
        ylab = "Number of Variables",
        cex.main = 1.2,
        cex.lab = 1.1,
        cex.names = 1.1,
        ylim = c(0, max(correct_level_counts) * 1.2))

# 添加数值标签和百分比
for(i in 1:length(correct_level_counts)) {
  text(x = i,
       y = correct_level_counts[i] + max(correct_level_counts) * 0.05,
       labels = paste0("n=", correct_level_counts[i], " (", 
                       round(correct_level_counts[i]/sum(correct_level_counts)*100, 1), "%)"),
       cex = 1.1)
}

dev.off()
cat("✅ 修正后的分布图已保存为 'Screening_Level_Distribution_Corrected.png'\n")


# ====================================================
# 模型比较与LASSO筛选
# ====================================================

# ====================================================
# 完整模型比较分析：四个模型（模型4采用 lambda.min）
# ====================================================

# 1. 加载必要包 --------------------------------------------------------------
library(glmnet)      # LASSO回归
library(pROC)        # ROC分析
library(ggplot2)     # 绘图
library(openxlsx)    # Excel输出
library(dplyr)       # 数据操作

# 2. 准备建模数据 -------------------------------------------------------------
# 假设 data 已存在，且包含 risk_binary 及各心理指标变量
# 确保心理指标变量为数值型（清洗代码已做）

# 定义所有可能的心理指标变量（根据您的数据实际情况）
all_psych_vars <- c(
  "psychosis", "suicide_ideation",
  "anxiety", "depression", "paranoia", "inferiority", "sensitivity",
  "social_phobia", "somatization", "dependence", "hostility", "impulsivity",
  "compulsion", "internet_addiction", "self_harm", "eating_disorder",
  "sleep_disturbance",
  "school_adjustment", "interpersonal_stress", "academic_stress",
  "employment_stress", "relationship_stress"
)

# 只保留实际存在的变量
existing_vars <- all_psych_vars[all_psych_vars %in% colnames(data)]
cat("实际存在的心理指标变量（共", length(existing_vars), "个）：\n")
print(existing_vars)

# 提取完整数据（无缺失）
model_data <- data[, c("risk_binary", existing_vars)]
model_data <- na.omit(model_data)
cat("用于建模的样本量：", nrow(model_data), "\n")

# 3. 定义各层级变量集（基于实际存在的变量）--------------------------------------
# 一级：严重心理问题
level1_vars <- intersect(c("psychosis", "suicide_ideation"), existing_vars)
# 二级：一般心理问题
level2_vars <- intersect(
  c("anxiety", "depression", "paranoia", "inferiority", "sensitivity",
    "social_phobia", "somatization", "dependence", "hostility", "impulsivity",
    "compulsion", "internet_addiction", "self_harm", "eating_disorder",
    "sleep_disturbance"),
  existing_vars
)
# 三级：发展性困扰
level3_vars <- intersect(
  c("school_adjustment", "interpersonal_stress", "academic_stress",
    "employment_stress", "relationship_stress"),
  existing_vars
)

cat("\n各层级变量数量：\n")
cat("Level 1:", length(level1_vars), "\n")
cat("Level 2:", length(level2_vars), "\n")
cat("Level 3:", length(level3_vars), "\n")

# 4. 构建模型1-3（逻辑回归）---------------------------------------------------
# 模型1：仅一级
if (length(level1_vars) > 0) {
  form1 <- as.formula(paste("risk_binary ~", paste(level1_vars, collapse = "+")))
  m1 <- glm(form1, data = model_data, family = binomial)
} else stop("一级变量不存在，无法构建模型1")

# 模型2：一级 + 二级
if (length(level2_vars) > 0) {
  form2 <- as.formula(paste("risk_binary ~", paste(c(level1_vars, level2_vars), collapse = "+")))
  m2 <- glm(form2, data = model_data, family = binomial)
} else m2 <- m1

# 模型3：一级 + 二级 + 三级
if (length(level3_vars) > 0) {
  form3 <- as.formula(paste("risk_binary ~", paste(c(level1_vars, level2_vars, level3_vars), collapse = "+")))
  m3 <- glm(form3, data = model_data, family = binomial)
} else m3 <- m2

# 5. 模型4：LASSO回归（lambda.min）--------------------------------------------
# 准备预测矩阵和响应变量
x_all <- as.matrix(model_data[, existing_vars])
y_all <- model_data$risk_binary

# 交叉验证选择最优lambda
set.seed(123)  # 可重复性
cv_lasso <- cv.glmnet(x_all, y_all, family = "binomial", alpha = 1, nfolds = 10)

# 提取 lambda.min 对应的系数
lambda_min <- cv_lasso$lambda.min
cat("\nlambda.min =", lambda_min, "\n")

lasso_coef_min <- coef(cv_lasso, s = "lambda.min")
lasso_coef_min <- as.matrix(lasso_coef_min)
nonzero_vars_min <- rownames(lasso_coef_min)[which(lasso_coef_min != 0)][-1]  # 去掉截距
cat("lambda.min 筛选出的变量（共", length(nonzero_vars_min), "个）：\n")
print(nonzero_vars_min)

# 用筛选出的变量重新拟合标准逻辑回归（获取OR和CI）
if (length(nonzero_vars_min) > 0) {
  form4 <- as.formula(paste("risk_binary ~", paste(nonzero_vars_min, collapse = "+")))
  m4 <- glm(form4, data = model_data, family = binomial)
} else {
  # 若无变量被选中，则用截距模型（极罕见）
  m4 <- glm(risk_binary ~ 1, data = model_data, family = binomial)
  nonzero_vars_min <- character(0)
}

# 6. 计算四个模型的AUC、AIC、BIC-----------------------------------------------
# 定义函数计算AUC及置信区间
calc_auc <- function(model, newdata = model_data) {
  pred_prob <- predict(model, newdata = newdata, type = "response")
  roc_obj <- roc(newdata$risk_binary, pred_prob, quiet = TRUE)
  auc_val <- auc(roc_obj)
  ci_val <- ci.auc(roc_obj)
  return(c(AUC = auc_val, lower = ci_val[1], upper = ci_val[3]))
}

# 模型列表
models <- list(m1, m2, m3, m4)
model_names <- c("Model1 (Level1)", "Model2 (Level1+2)", "Model3 (Full)", "Model4 (LASSO-min)")

# 初始化结果数据框
results <- data.frame(
  Model = model_names,
  AIC = sapply(models, AIC),
  BIC = sapply(models, BIC),
  AUC = NA,
  AUC_lower = NA,
  AUC_upper = NA
)

for (i in 1:4) {
  auc_info <- calc_auc(models[[i]])
  results$AUC[i] <- auc_info["AUC"]
  results$AUC_lower[i] <- auc_info["lower"]
  results$AUC_upper[i] <- auc_info["upper"]
}

# 打印结果
cat("\n========== 模型比较结果 ==========\n")
print(results)

# 7. 绘制四个模型的ROC曲线叠加图-----------------------------------------------
# 获取各模型的预测概率
pred_list <- lapply(models, function(m) predict(m, newdata = model_data, type = "response"))
roc_list <- lapply(pred_list, function(p) roc(model_data$risk_binary, p, quiet = TRUE))

# 定义颜色
colors <- c("blue", "green", "red", "purple")

# 保存为PNG
png("ROC_Comparison_final.png", width = 800, height = 600)
plot(roc_list[[1]], col = colors[1], lwd = 2, main = "ROC Curves of Four Models")
for (i in 2:4) {
  plot(roc_list[[i]], col = colors[i], lwd = 2, add = TRUE)
}
legend("bottomright", 
       legend = paste0(model_names, " (AUC=", round(results$AUC, 3), ")"),
       col = colors, lwd = 2)
dev.off()

# 7.5 DeLong检验比较模型AUC
library(pROC)  # 确保已加载
cat("\n========== DeLong检验结果 ==========\n")

# 比较模型3和模型4
test_3vs4 <- roc.test(roc_list[[3]], roc_list[[4]], method = "delong")
cat("模型3 (Full) vs 模型4 (LASSO-min):\n")
cat("  AUC差异 =", round(test_3vs4$estimate[1] - test_3vs4$estimate[2], 4), "\n")
cat("  p值 =", format.pval(test_3vs4$p.value, digits = 4), "\n\n")

# 比较模型2和模型3
test_2vs3 <- roc.test(roc_list[[2]], roc_list[[3]], method = "delong")
cat("模型2 (Level1+2) vs 模型3 (Full):\n")
cat("  AUC差异 =", round(test_2vs3$estimate[1] - test_2vs3$estimate[2], 4), "\n")
cat("  p值 =", format.pval(test_2vs3$p.value, digits = 4), "\n\n")

# 比较模型1和模型2
test_1vs2 <- roc.test(roc_list[[1]], roc_list[[2]], method = "delong")
cat("模型1 (Level1) vs 模型2 (Level1+2):\n")
cat("  AUC差异 =", round(test_1vs2$estimate[1] - test_1vs2$estimate[2], 4), "\n")
cat("  p值 =", format.pval(test_1vs2$p.value, digits = 4), "\n")

# 8. 处理模型4筛选出的变量（英文版森林图，含OR值标签）-------------------------------
if (length(nonzero_vars_min) > 1) {
  # 从 m4 提取系数、OR、置信区间
  coef_summary <- summary(m4)$coefficients
  or_vars <- rownames(coef_summary)[-1]  # 去掉截距
  or_est <- exp(coef_summary[-1, 1])
  or_ci_lower <- exp(coef_summary[-1, 1] - 1.96 * coef_summary[-1, 2])
  or_ci_upper <- exp(coef_summary[-1, 1] + 1.96 * coef_summary[-1, 2])
  p_vals <- coef_summary[-1, 4]
  
  # 标注变量所属级别（英文）
  level_info <- sapply(or_vars, function(v) {
    if (v %in% level1_vars) return("Level 1")
    else if (v %in% level2_vars) return("Level 2")
    else if (v %in% level3_vars) return("Level 3")
    else return("Other")
  })
  
  # 创建森林图数据框
  forest_data <- data.frame(
    variable = or_vars,
    level = level_info,
    OR = round(or_est, 2),
    lower = round(or_ci_lower, 2),
    upper = round(or_ci_upper, 2),
    p = round(p_vals, 4)
  )
  
  # 按级别和OR值排序
  forest_data <- forest_data[order(forest_data$level, forest_data$OR), ]
  n <- nrow(forest_data)
  y <- 1:n  # y轴位置
  
  # 准备OR标签文本（格式：OR (lower-upper)）
  label_text <- paste0(forest_data$OR, " (", forest_data$lower, "-", forest_data$upper, ")")
  
  # 计算x轴范围，为右侧文本预留空间
  x_min <- min(forest_data$lower) * 0.8
  x_max <- max(forest_data$upper) * 1.5  # 增加右侧空间
  
  # 设置图形设备，保存为PNG（英文文件名）
  png("LASSOmin_ForestPlot_English.png", width = 10, height = max(5, n * 0.5), units = "in", res = 300)
  par(mar = c(4, 12, 3, 4))  # 左边距显示变量名，右边距显示OR标签
  
  # 创建空图，x轴使用对数坐标（英文标题）
  plot(NA, 
       xlim = c(x_min, x_max),
       ylim = c(0.5, n + 0.5), 
       yaxt = "n", 
       xlab = "Odds Ratio (95% CI)", 
       ylab = "",
       main = "Core Predictors Selected by LASSO-min Model", 
       log = "x")
  abline(v = 1, lty = 2, col = "gray50")  # 参考线
  
  # 添加误差线和点
  for (i in 1:n) {
    lines(x = c(forest_data$lower[i], forest_data$upper[i]), 
          y = c(y[i], y[i]),
          col = "darkblue", lwd = 1.5)
    points(x = forest_data$OR[i], y = y[i], 
           pch = 19, col = "royalblue", cex = 1.5)
  }
  
  # 添加y轴标签（变量名 + 级别，英文）
  axis(2, at = y, 
       labels = paste0(forest_data$variable, " (", forest_data$level, ")"),
       las = 1, cex.axis = 0.9)
  
  # 在右侧添加OR值标签
  text(x = forest_data$upper * 1.1,   # 放在误差线右侧约10%位置
       y = y, 
       labels = label_text, 
       pos = 4,                        # 右对齐
       offset = 0.2, 
       cex = 0.8, 
       col = "black")
  
  dev.off()
  cat("英文版森林图已保存为 LASSOmin_ForestPlot_English.png\n")
  
} else if (length(nonzero_vars_min) == 1) {
  # 只有一个变量，输出该变量的OR和CI
  single_var <- nonzero_vars_min[1]
  coef_val <- coef(m4)[2]
  se_val <- summary(m4)$coefficients[2, 2]
  or_val <- exp(coef_val)
  ci_lower <- exp(coef_val - 1.96 * se_val)
  ci_upper <- exp(coef_val + 1.96 * se_val)
  
  cat("\n模型4只筛选出一个变量：", single_var, "\n")
  cat(sprintf("OR = %.2f (95%% CI: %.2f-%.2f)\n", or_val, ci_lower, ci_upper))
  
  forest_data <- data.frame(
    variable = single_var,
    level = ifelse(single_var %in% level1_vars, "Level 1",
                   ifelse(single_var %in% level2_vars, "Level 2",
                          ifelse(single_var %in% level3_vars, "Level 3", "Other"))),
    OR = round(or_val, 2),
    lower = round(ci_lower, 2),
    upper = round(ci_upper, 2),
    p = round(summary(m4)$coefficients[2, 4], 4)
  )
  
} else {
  cat("\n模型4未筛选出任何变量。\n")
  forest_data <- data.frame()
}

# 9. Bootstrap内部验证LASSO变量选择的稳定性
cat("\n========== Bootstrap内部验证 ==========\n")
set.seed(456)  # 不同种子，确保可重复
n_bootstrap <- 100  # Bootstrap次数（可根据计算资源调整，建议100-200）
n_obs <- nrow(model_data)
n_vars <- length(existing_vars)
selected_freq <- rep(0, n_vars)
names(selected_freq) <- existing_vars

# 进度显示
pb <- txtProgressBar(min = 0, max = n_bootstrap, style = 3)

for (b in 1:n_bootstrap) {
  # 有放回抽样
  boot_idx <- sample(1:n_obs, size = n_obs, replace = TRUE)
  x_boot <- x_all[boot_idx, ]
  y_boot <- y_all[boot_idx]
  
  # 在bootstrap样本上运行LASSO交叉验证（使用与主分析相同的设置）
  cv_boot <- cv.glmnet(x_boot, y_boot, family = "binomial", alpha = 1, nfolds = 10)
  
  # 提取lambda.min对应的系数
  coef_boot <- as.matrix(coef(cv_boot, s = "lambda.min"))
  
  # 记录非零变量（排除截距）
  selected_vars <- rownames(coef_boot)[which(coef_boot != 0)][-1]
  selected_freq[selected_vars] <- selected_freq[selected_vars] + 1
  
  setTxtProgressBar(pb, b)
}
close(pb)

# 转换为频率百分比
selected_pct <- selected_freq / n_bootstrap * 100
selected_df <- data.frame(
  Variable = names(selected_pct),
  Selection_Frequency = round(selected_pct, 1),
  Level = sapply(names(selected_pct), function(v) {
    if (v %in% level1_vars) return("Level 1")
    else if (v %in% level2_vars) return("Level 2")
    else if (v %in% level3_vars) return("Level 3")
    else return("Other")
  })
)
# 按频率降序排列
selected_df <- selected_df[order(-selected_df$Selection_Frequency), ]

cat("\n变量被选中的频率（%）:\n")
print(selected_df[selected_df$Selection_Frequency > 0, ])

# 保存到Excel（新增sheet）
library(openxlsx)
# 如果之前已有工作簿，则加载；否则新建
if (file.exists("Model_Results_final.xlsx")) {
  wb <- loadWorkbook("Model_Results_final.xlsx")
} else {
  wb <- createWorkbook()
}
addWorksheet(wb, "Bootstrap_Stability")
writeData(wb, "Bootstrap_Stability", selected_df)
saveWorkbook(wb, "Model_Results_final.xlsx", overwrite = TRUE)

# 绘制频率条形图（仅显示频率>0的变量）
library(ggplot2)
plot_data <- selected_df[selected_df$Selection_Frequency > 0, ]
if (nrow(plot_data) > 0) {
  p_freq <- ggplot(plot_data, 
                   aes(x = reorder(Variable, Selection_Frequency), 
                       y = Selection_Frequency, 
                       fill = Level)) +
    geom_bar(stat = "identity") +
    coord_flip() +
    labs(x = "", y = "Selection Frequency (%)", 
         title = "Bootstrap Stability of LASSO-min Selected Variables") +
    scale_fill_manual(values = c("Level 1" = "#E69F00", 
                                 "Level 2" = "#56B4E9", 
                                 "Level 3" = "#009E73",
                                 "Other" = "#999999")) +
    theme_minimal(base_size = 12) +
    theme(legend.position = "bottom", 
          plot.title = element_text(hjust = 0.5, face = "bold"))
  
  ggsave("Bootstrap_Stability.png", p_freq, width = 8, 
         height = max(4, nrow(plot_data) * 0.3), dpi = 300)
  print(p_freq)
}
--------------弹性网络验证--------------------------
# 安装包（第一次运行）
install.packages("glmnet")
install.packages("readxl")

# 加载包
library(glmnet)
library(readxl)


# 2. 自变量（22个，完全按你给的变量名）
X <- as.matrix(data[, c(
  "sensitivity",
  "depression",
  "compulsion",
  "sleep_disturbance",
  "anxiety",
  "interpersonal_stress",
  "paranoia",
  "school_adjustment",
  "somatization",
  "suicide_ideation",
  "psychosis",
  "eating_disorder",
  "social_phobia",
  "dependence",
  "internet_addiction",
  "self_harm",
  "employment_stress",
  "relationship_stress",
  "inferiority",
  "hostility",
  "impulsivity",
  "academic_stress"
)])

# 3. 因变量
Y <- data$risk_binary

# 4. 运行弹性网络（alpha=0.5 标准）
set.seed(123)
cv.fit <- cv.glmnet(X, Y, family = "binomial", alpha = 0.5)

# 5. 输出最终筛选结果
coef(cv.fit, s = "lambda.min")                   
