#define _CRT_SECURE_NO_WARNINGS 1
#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <windows.h>
#include <shlobj.h>
#include <commdlg.h>  // 需要包含这个头文件

#pragma comment(lib, "comdlg32.lib")  // 链接库

typedef struct Project {
    int id;
    char name[20];
    char main_person[20];
    char start_date[20];
    char end_date[20];
    float process;
}Pro, * PPro;

int project_nums = 0;//数组索引
Pro projects[100];

// 初始化菜单
void initMenu() {
    printf("===========欢迎使用企业项目任务跟踪系统==========\n");
    printf("1.录入项目\n");
    printf("2.修改项目\n");
    printf("3.查询项目\n");
    printf("4.统计项目\n");
    printf("5.排序并显示所有项目\n");
    printf("6.删除项目\n");
    printf("7.保存项目\n");
    printf("8.读取文件\n");
    printf("选择要执行的功能\n");
}

// 录入项目
void typeIn(PPro projects) {
    Pro pro;
    printf("录入项目编号\n");
    scanf("%d", &pro.id);
    printf("录入项目名称\n");
    scanf("%s", &pro.name);
    printf("录入项目负责人\n");
    scanf("%s", &pro.main_person);
    printf("录入项目开始日期，形式要求xxxx-xx-xx\n");
    scanf("%s", &pro.start_date);
    printf("录入项目计划完成日期，形式要求xxxx-xx-xx\n");
    scanf("%s", &pro.end_date);
    printf("录入项目进度百分比\n");
    scanf("%f", &pro.process);

    projects[project_nums] = pro;
    ++project_nums;
    printf("项目录入完成！\n");
}

// 显示项目信息
void printInfo(int number) {
    printf("项目ID: %d, 项目名称: %s, 负责人: %s, 开始日期: %s, 结束日期: %s, 进度: %f%%\n\n",
        projects[number].id, projects[number].name, projects[number].main_person, projects[number].start_date,
        projects[number].end_date, projects[number].process);
}

// 根据id查找找项目
int searchProWithId() {
    int id_;
    printf("输入项目id\n");
    scanf("%d", &id_);
    for (int i = 0; i < project_nums; i++) {
        if (id_ == projects[i].id) {
            printInfo(i);
            return i;
        }
    }
    printf("不存在该项目，请输入正确id!\n");
    return -1;
}

// 更新项目进度与负责人
void reviseProject() {
    int res = searchProWithId();
    if (res == -1) return;
    printf("输入新的项目进度\n");
    scanf("%f", &projects[res].process);
    printf("输入新的项目负责人\n");
    scanf("%s", &projects[res].main_person);
    printf("修改成功！\n");
}

// 获取当前日期
char* getCurrentDate() {
    time_t now = time(NULL);
    struct tm* tm_now = localtime(&now);

    int year = tm_now->tm_year + 1900; // tm_year 存储的是从1900年开始的年数
    int month = tm_now->tm_mon + 1; // tm_mon 存储的是 0-11
    int day = tm_now->tm_mday;
    char str[20];
    sprintf(str, "%d-%02d-%02d", year, month, day);

    return str;
}

// 统计项目完成率与延期数量
void calFinishedRateAndTimeoutNums() {
    int finished_projects = 0;
    int timeout_projects = 0;
    for (int i = 0; i < project_nums; i++) {
        if ((int)projects[i].process == 100) ++finished_projects;
        if (projects[i].end_date < getCurrentDate()) ++timeout_projects;
    }
    printf("项目完成率：%f\n", (float)finished_projects / (float)project_nums);
    printf("延期项目数：%d\n", timeout_projects);
}

// 按进度进行排序
int compareByProcessDesc(const void* a, const void* b) {
    Pro* proA = (Pro*)a;
    Pro* proB = (Pro*)b;

    if (proA->process > proB->process) return -1;
    if (proA->process < proB->process) return 1;
    return 0;  // 相等
}

void sortAndShow() {
    qsort(projects, project_nums, sizeof(Pro), compareByProcessDesc);
    for (int i = 0; i < project_nums; i++) printInfo(i);
    return;
}

// 根据id删除项目
void deleteProject() {
    int number = searchProWithId();
    if (number == -1) return;
    for (int i = number; i < project_nums - 1; i++) {
        projects[i] = projects[i + 1];
    }
    printf("删除成功！\n");
}

// 保存为格式化文本文件
void saveAsFormattedText() {
    char filepath[MAX_PATH];
    SHGetFolderPathA(NULL, CSIDL_DESKTOP, NULL, 0, filepath);
    strcat(filepath, "\\projects.txt");

    FILE* fp = fopen(filepath, "w");
    if (fp == NULL) {
        printf("无法创建文本文件！\n");
        return;
    }

    // 写入标题
    fprintf(fp, "========== 项目管理列表 ==========\n");
    fprintf(fp, "生成时间：%s", __DATE__);
    fprintf(fp, " %s\n\n", __TIME__);

    // 写入表头
    fprintf(fp, "┌──────┬──────────────────────┬──────────┬────────────┬────────────┬────────┐\n");
    fprintf(fp, "│ %-4s │ %-20s │ %-8s │ %-10s │ %-10s │ %-6s │\n",
        "ID", "项目名称", "负责人", "开始日期", "结束日期", "进度%%");
    fprintf(fp, "├──────┼──────────────────────┼──────────┼────────────┼────────────┼────────┤\n");

    // 写入数据
    for (int i = 0; i < project_nums; i++) {
        fprintf(fp, "│ %-4d │ %-20s │ %-8s │ %-10s │ %-10s │ %6.2f │\n",
            projects[i].id,
            projects[i].name,
            projects[i].main_person,
            projects[i].start_date,
            projects[i].end_date,
            projects[i].process);
    }

    fprintf(fp, "└──────┴──────────────────────┴──────────┴────────────┴────────────┴────────┘\n");

    // 写入统计信息
    fprintf(fp, "\n总计：%d 个项目\n", project_nums);

    // 计算平均进度
    float total_process = 0;
    for (int i = 0; i < project_nums; i++) {
        total_process += projects[i].process;
    }
    fprintf(fp, "平均进度：%.2f%%\n", total_process / project_nums);

    fclose(fp);
    printf("格式化文本文件已保存到桌面：%s\n", filepath);
}

// 打开文件对话框
int openFileDialog(char* filename, int size) {
    OPENFILENAMEA ofn;
    char fileBuffer[MAX_PATH] = { 0 };

    // 初始化结构体
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = NULL;  // 父窗口句柄
    ofn.lpstrFile = fileBuffer;
    ofn.nMaxFile = sizeof(fileBuffer);

    // 设置文件过滤器（只显示txt和csv文件）
    ofn.lpstrFilter = "所有支持文件\0*.txt;*.csv\0文本文件\0*.txt\0CSV文件\0*.csv\0所有文件\0*.*\0";
    ofn.nFilterIndex = 1;  // 默认使用第一个过滤器

    ofn.lpstrTitle = "选择项目数据文件";  // 对话框标题
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;

    // 打开对话框
    if (GetOpenFileNameA(&ofn)) {
        strncpy(filename, fileBuffer, size);
        return 1;  // 成功选择文件
    }

    return 0;  // 用户取消
}

// 保存文件对话框
int saveFileDialog(char* filename, int size) {
    OPENFILENAMEA ofn;
    char fileBuffer[MAX_PATH] = { 0 };

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.lpstrFile = fileBuffer;
    ofn.nMaxFile = sizeof(fileBuffer);
    ofn.lpstrFilter = "CSV文件\0*.csv\0文本文件\0*.txt\0所有文件\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.lpstrTitle = "保存项目数据";
    ofn.lpstrDefExt = "csv";  // 默认扩展名
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;

    if (GetSaveFileNameA(&ofn)) {
        strncpy(filename, fileBuffer, size);
        return 1;
    }
    return 0;
}

// 从文件加载项目
void loadProjectsFromFile() {
    char filepath[MAX_PATH];

    if (!openFileDialog(filepath, sizeof(filepath))) {
        printf("操作已取消\n");
        return;
    }

    FILE* fp = fopen(filepath, "r");
    if (fp == NULL) {
        printf("无法打开文件：%s\n", filepath);
        return;
    }

    // 读取并跳过表头
    char buffer[256];
    fgets(buffer, sizeof(buffer), fp);

    // 读取数据
    project_nums = 0;
    while (fscanf(fp, "%d,%[^,],%[^,],%[^,],%[^,],%f\n",
        &projects[project_nums].id,
        projects[project_nums].name,
        projects[project_nums].main_person,
        projects[project_nums].start_date,
        projects[project_nums].end_date,
        &projects[project_nums].process) == 6) {
        project_nums++;
        if (project_nums >= 100) break;  // 防止数组越界
    }

    fclose(fp);
    printf("成功从文件加载 %d 个项目\n", project_nums);
    printf("文件位置：%s\n", filepath);
}

int main()
{
    initMenu();
    int choice = 0;
    do {
        scanf("%d", &choice);
        switch (choice) {
        case 1:
            //录入
            typeIn(projects);
            break;
        case 2:
            // 修改
            reviseProject();
            break;
        case 3:
            //查询
            searchProWithId();
            break;
        case 4:
            //统计项目完成率和延期数量
            calFinishedRateAndTimeoutNums();
            break;
        case 5:
            // 按进度进行排序
            sortAndShow();
            break;
        case 6:
            // 删除项目
            deleteProject();
            break;
        case 7:
            // 文件保存
            saveAsFormattedText();
            break;
        case 8:
            // 文件
            loadProjectsFromFile();
            break;
        default:
            printf("输入错误，请重新选择\n");
        }
    } while (1);
    return 0;
}
