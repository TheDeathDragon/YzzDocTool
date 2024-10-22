#include <iostream>
#include <filesystem>
#include <string>
#include <sstream>
#include <iomanip>
#include <ctime>
#include <windows.h>
#define _SILENCE_EXPERIMENTAL_FILESYSTEM_DEPRECATION_WARNING
#include <experimental/filesystem>
namespace fs = std::experimental::filesystem;

std::tm getStartDate(int year, int month) {
    std::tm startDate = {};
    startDate.tm_year = year - 1900; // Year since 1900
    startDate.tm_mon = month - 1; // Month range [0, 11]
    startDate.tm_mday = 1; // Start at day 1
    mktime(&startDate); // Normalize the time structure
    return startDate;
}

std::tm addDays(const std::tm& date, int days) {
    std::tm newDate = date;
    newDate.tm_mday += days;
    mktime(&newDate); // Normalize the time structure
    return newDate;
}

std::string formatDate(const std::tm& date, const std::string& tableName) {
    std::ostringstream oss;
    oss << std::put_time(&date, "%Y年%m月%d日") << tableName << ".xlsx";
    return oss.str();
}

int main() {
    int year, month;
    std::string tableName;

    // Prompt user for input
    std::cout << "请输入年份: ";
    std::cin >> year;
    std::cout << "请输入月份: ";
    std::cin >> month;
    std::cout << "请输入表格名字: ";
    std::cin >> tableName;

    // Get the current directory
    fs::path currentDirectory = fs::current_path();

    // Define template file path and destination folder
    fs::path templateFilePath = currentDirectory / "empty.xlsx";
    std::tm startDate = getStartDate(year, month);
    std::tm endDate = addDays(startDate, 29);

    std::ostringstream folderName;
    folderName << std::put_time(&startDate, "%Y_%m");
    fs::path destinationFolder = currentDirectory / folderName.str();

    // Create destination folder if it doesn't exist
    if (!fs::exists(destinationFolder)) {
        fs::create_directory(destinationFolder);
    }

    std::tm currentDate = startDate;
    while (std::difftime(std::mktime(&currentDate), std::mktime(&endDate)) <= 0) {
        // Create new file name
        std::string newFileName = formatDate(currentDate, tableName);
        fs::path newFilePath = destinationFolder / newFileName;

        // Copy the template file
        try {
            fs::copy_file(templateFilePath, newFilePath, fs::copy_options::overwrite_existing);
        }
        catch (fs::filesystem_error& e) {
            std::cerr << "Error copying file: " << e.what() << '\n';
            return 1;
        }

        // Add one day to the current date
        currentDate = addDays(currentDate, 1);
    }

    std::cout << "文件生成完成, 按任意键退出！\n";
    getchar();
    return 0;
}
