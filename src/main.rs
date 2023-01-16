use std::path::{Path, PathBuf};

use anyhow::Result;
use calamine::{open_workbook, DataType, Reader, Xlsx};
use console::style;
use dialoguer::Input;
use indicatif::{ProgressBar, ProgressStyle};
use rfd::FileDialog;
use xlsxwriter::{Workbook, Worksheet};

fn main() -> Result<()> {
    let file_path = open_file()?;
    let save_path = get_save_path()?;

    let count: usize = Input::new()
        .with_prompt("请输入每个文件行数")
        .default(300)
        .interact_text()?;

    let title_rows: usize = Input::new()
        .with_prompt("标题行(输入0不保留标题)")
        .default(1)
        .interact_text()?;

    // 进度条样式
    let spinner_style = ProgressStyle::with_template(
        "{prefix:.bold.dim} {bar:40} 正在保存第{pos}/{len}个文件 {msg}",
    )
    .unwrap()
    .tick_chars("⠁⠂⠄⡀⢀⠠⠐⠈ ");

    println!("{} 正在读取文件...", style("[1/3]").bold().dim());

    let mut excel: Xlsx<_> = open_workbook(&file_path)?;

    println!("{} 正在解析表格...", style("[2/3]").bold().dim());
    // 读取表格第一个sheet
    if let Some(Ok(r)) = excel.worksheet_range_at(0) {
        // 会生成的总文件数
        let lines = r.rows().len() / count;

        // 显示进度条，设置位置为1
        let pb = ProgressBar::new((lines + 1) as u64)
            .with_style(spinner_style)
            .with_position(1);

        println!("{} 正在保存文件...", style("[3/3]").bold().dim());

        let mut row_index = 0u32;
        let mut page = 1usize;

        // 创建表格
        let mut workbook = crate_xlsx(&save_path, &file_path, page)?;
        let mut sheet1 = workbook.add_worksheet(Some("Sheet1"))?;

        // 获取标题
        let title: Vec<&[DataType]> = r.rows().take(title_rows).collect();

        for (i, row) in r.rows().enumerate() {
            if i == count * page {
                // 保存并创建新的表格
                workbook.close()?;
                pb.inc(1);
                page += 1;
                row_index = 0;
                workbook = crate_xlsx(&save_path, &file_path, page)?;
                sheet1 = workbook.add_worksheet(Some("Sheet1"))?;

                // 写入标题
                for tr in title.clone() {
                    for (i, data) in tr.iter().enumerate() {
                        write_sheet(&mut sheet1, data, i, row_index)?;
                    }
                    row_index += 1;
                }
            }

            // 判断类型 写入表格
            for (ii, data) in row.iter().enumerate() {
                write_sheet(&mut sheet1, data, ii, row_index)?;
            }
            row_index += 1;
        }
        pb.finish_with_message("完成！");
    }
    Ok(())
}

/// 创建Excel
fn crate_xlsx(save_path: &Path, file_name: &Path, page: usize) -> Result<Workbook> {
    let workbook = Workbook::new(
        save_path
            .join(format!(
                "{}-{}.{}",
                file_name.file_stem().unwrap().to_str().unwrap(),
                page,
                file_name.extension().unwrap().to_str().unwrap()
            ))
            .to_str()
            .unwrap(),
    )?;
    Ok(workbook)
}

fn open_file() -> Result<PathBuf> {
    Ok(FileDialog::new()
        .add_filter("Excel 文件", &["xlsx", "xls"])
        .set_directory("~/Desktop")
        .set_title("选择要拆分的Excel文件")
        .pick_file()
        .unwrap())
}
fn get_save_path() -> Result<PathBuf> {
    Ok(FileDialog::new()
        .set_title("保存到")
        .set_directory("~/Desktop")
        .pick_folder()
        .unwrap())
}

fn write_sheet(sheet1: &mut Worksheet, data: &DataType, ii: usize, row_index: u32) -> Result<()> {
    if data.is_bool() {
        sheet1.write_boolean(row_index, ii as u16, data.get_bool().unwrap_or(false), None)?;
    }
    if data.is_empty() {
        sheet1.write_string(row_index, ii as u16, "", None)?;
    }
    if data.is_float() || data.is_int() {
        sheet1.write_number(row_index, ii as u16, data.get_float().unwrap_or(0f64), None)?;
    }
    if data.is_string() {
        sheet1.write_string(row_index, ii as u16, data.get_string().unwrap_or(""), None)?;
    }
    Ok(())
}
