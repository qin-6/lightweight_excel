// profile.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include <iostream>
#include <string>
#include<fstream>
#include<vector>
#include<sstream>
#include<io.h>
#include"BasicExcel.hpp"
using namespace std;
using namespace YExcel;
string pname[7] =
{
   "server_bitrate_send:",
   "frame_delay:",
   "render_time:",
   "avg_dec_time:",
   "avg_enc_time:",
   "rtt:",
   "video_AvgSendDelay "
};
vector<uint32_t> fff(string &str, string &substr, float &avg)
{
    vector<uint32_t> vec;
    uint32_t len = substr.length();
    uint32_t pos = 0;
    uint32_t last_pos = str.rfind(substr);
    uint32_t sum = 0;
    uint32_t count = 0;
    while (true)
    {
        pos = str.find(substr, pos) + len;
        uint32_t pos2 = str.find_first_of('\n', pos);
        string value = str.substr(pos, pos2 - pos);
        cout << value << "\n";
        uint32_t v = stoi(value);
        sum += v;
        vec.push_back(v);
        count ++;
        if (pos > last_pos)
        {
            break;
        }
    }
    avg = static_cast<float>(sum / count);
    return vec;
}
void getFileNames(string path, vector<string>& files)
{
    //文件句柄
    //注意：我发现有些文章代码此处是long类型，实测运行中会报错访问异常
    intptr_t hFile = 0;
    //文件信息
    struct _finddata_t fileinfo;
    string p;
    if ((hFile = _findfirst(p.assign(path).append("\\*").c_str(), &fileinfo)) != -1)
    {
        do
        {
            //如果是目录,递归查找
            //如果不是,把文件绝对路径存入vector中
            if ((fileinfo.attrib & _A_SUBDIR))
            {
                if (strcmp(fileinfo.name, ".") != 0 && strcmp(fileinfo.name, "..") != 0)
                    getFileNames(p.assign(path).append("\\").append(fileinfo.name), files);
            }
            else
            {
                files.push_back(p.assign(path).append("\\").append(fileinfo.name));
            }
        } while (_findnext(hFile, &fileinfo) == 0);
        _findclose(hFile);
    }
}
vector<uint32_t> Parse_file(string &name, BasicExcelWorksheet* sheet)
{
    vector<uint32_t> profile_avg;
    string  file_name;
    string  file_path;
    uint32_t pos1 = name.find_last_of("\\") + 1;
    uint32_t pos2 = name.find_last_of(".");
    file_name = name.substr(pos1, pos2 - pos1);
    file_path = name.substr(0, pos1);
    ifstream profile_file;
    profile_file.open(name, ios::in);
    if (profile_file.is_open())
    {
        BasicExcelCell* cell;
        stringstream ss;
        uint32_t pos = 0;
        ss << profile_file.rdbuf();
        string text = ss.str();
        cout << "Input file Length is " << text.length() / 1024 << "KB\n";
        for (int i = 0; i < 7; i++)
        {
            float avg;
            vector<uint32_t> vec;
            vec = fff(text, pname[i], avg);
            vec.push_back(static_cast<uint32_t>(avg));
            cell = sheet->Cell(0, i + 1);
            cell->SetString(pname[i].c_str());
            cell = sheet->Cell(vec.size(), 0);
            cell->SetString("average");
            for (int j = 0; j < vec.size(); j++)
            {
                cell = sheet->Cell(j + 1, i + 1);
                cell->SetInteger((int)vec[j]);
            }
            profile_avg.push_back(static_cast<uint32_t>(avg));
        }
    }
    return profile_avg;
}
int main()
{
    BasicExcel e;
    BasicExcelCell* cell;
    string file;
    string file_name;
    string file_path;
    vector<string> files;
    vector<string> files_filtered;
    cout << "Please Input the Profile file:\n";
    getline(cin, file_path, '\n');
    getFileNames(file_path, files);
    for (auto name : files)
    {
        uint32_t pos = name.find_last_of(".") + 1;
        string extral_name = name.substr(pos, name.length() - pos);
        if (extral_name == "txt" /* || extral_name == "log"*/)
        {
            files_filtered.push_back(name);
        }
    }
    e.New(files_filtered.size() + 1); 
    uint32_t col = 0;
    vector<vector<uint32_t>> vec;
    vector<string> item_names;
    for (auto name : files_filtered)
    {
        cout << name << endl;
        uint32_t pos1 = name.find_last_of(".");
        uint32_t pos2 = name.find_last_of("\\") + 1;
        string item_name = name.substr(pos2, pos1 - pos2);
        BasicExcelWorksheet* sheet = e.GetWorksheet((unsigned)col);
        sheet->Rename(item_name.c_str());
        item_names.push_back(item_name);
        vec.push_back(Parse_file(name, sheet));  
        col++;
    }
    BasicExcelWorksheet* sheet = e.GetWorksheet((unsigned)col);
    sheet->Rename("total");
    for (int i = 0; i < 7; i++)
    {
        cell = sheet->Cell(0, i + 1);
        cell->SetString(pname[i].c_str());
    }
    for (int i = 0; i < vec.size(); i++)
    {
        cell = sheet->Cell(i + 1, 0);
        cell->SetString(item_names[i].c_str());
        for(int j = 0; j < vec[i].size(); j ++)
        {    
            cell = sheet->Cell(i + 1, j + 1);
            cell->SetInteger((int)vec[i][j]);
        }
    }

    string file_new = file_path + "\\total.xls";
    e.SaveAs(file_new.c_str());
    while (true);
}

// 运行程序: Ctrl + F5 或调试 >“开始执行(不调试)”菜单
// 调试程序: F5 或调试 >“开始调试”菜单

// 入门使用技巧: 
//   1. 使用解决方案资源管理器窗口添加/管理文件
//   2. 使用团队资源管理器窗口连接到源代码管理
//   3. 使用输出窗口查看生成输出和其他消息
//   4. 使用错误列表窗口查看错误
//   5. 转到“项目”>“添加新项”以创建新的代码文件，或转到“项目”>“添加现有项”以将现有代码文件添加到项目
//   6. 将来，若要再次打开此项目，请转到“文件”>“打开”>“项目”并选择 .sln 文件
