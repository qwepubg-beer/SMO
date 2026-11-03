using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace WpfApp1.Pages
{
    /// <summary>
    /// Логика взаимодействия для Page1.xaml
    /// </summary>
    public partial class Page1 : Page
    {
        public Page1()
        {
            InitializeComponent();
        }

        private void Choose_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (QueueLimitation == null || QueueWithoutlimitation == null) return;
            QueueLimitation.Visibility = Visibility.Collapsed;
            QueueWithoutlimitation.Visibility = Visibility.Collapsed;
            switch (Choose.SelectedIndex)
            {
                case 1:
                    QueueLimitation.Visibility = Visibility.Visible;
                    break;
                case 2:
                    QueueWithoutlimitation.Visibility = Visibility.Visible;
                    break;
            }
        }

        private void Answer_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(IPZ.Text) || string.IsNullOrWhiteSpace(IPO.Text))
                {
                    MessageBox.Show("Заполните значения интенсивности потока заявок и обслуживания.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 1. Обработка Lambda (λ)
                double lambda = double.Parse(IPZ.Text);
                if (ChooseTimeIPZ.SelectedIndex == 1) // Если выбраны часы
                {
                    lambda = lambda / 60.0; // переводим в минуты
                }

                // 2. Обработка Mu (μ) или T_obs (t_обс)
                double inputValue = double.Parse(IPO.Text);
                double mu, t_obs;

                // Проверяем, что выбрал пользователь: 0 - это μ, 1 - это t_обс
                if (ServiceType.SelectedIndex == 0)
                {
                    // Пользователь ввел μ (интенсивность)
                    mu = inputValue;
                    if (Tops.SelectedIndex == 1) mu = mu / 60.0; // переводим заявки/час в заявки/мин
                    t_obs = 1.0 / mu; // вычисляем время
                }
                else
                {
                    // Пользователь ввел t_обс (время)
                    t_obs = inputValue;
                    if (Tops.SelectedIndex == 1) t_obs = t_obs * 60.0; // переводим часы в минуты
                    mu = 1.0 / t_obs; // вычисляем интенсивность
                }

                // Интенсивность нагрузки
                double rho = lambda * t_obs;

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Системы массового обслуживания.");
                sb.AppendLine($"Интенсивность потока обслуживания: μ = {mu:F3} заявок/мин.");
                sb.AppendLine($"Среднее время обслуживания: t_обс = 1/μ = {t_obs:F3} мин.");

                switch (Choose.SelectedIndex)
                {
                    case 0: // С отказами
                        sb.AppendLine($"\nИнтенсивность нагрузки.\nρ = λ * t_обс = {lambda:F3} * {t_obs:F3} = {rho:F3}");
                        double p0 = 1.0 / (1.0 + rho);
                        sb.AppendLine($"Вероятность, что канал свободен (доля времени простоя канала).\np0 = 1 / (1 + ρ) = 1 / (1 + {rho:F3}) = {p0:F3}");

                        double p_otk = 1.0 - p0;
                        sb.AppendLine($"Доля заявок, получивших отказ (вероятность отказа).\np_отк = 1 - p0 = 1 - {p0:F3} = {p_otk:F3}");

                        double Q = p0;
                        sb.AppendLine($"Относительная пропускная способность.\nДоля обслуживаемых заявок, поступающих в единицу времени:\nQ = p0 = {Q:F3}");

                        double A = Q * lambda;
                        sb.AppendLine($"Абсолютная пропускная способность.\nA = Q * λ = {Q:F3} * {lambda:F3} = {A:F3} заявок/мин.");

                        double t_pr = p_otk * t_obs;
                        sb.AppendLine($"Среднее время простоя СМО (мин.).\ntпр = p_отк * t_обс = {p_otk:F3} * {t_obs:F3} = {t_pr:F3}");

                        double L_obs = rho * Q;
                        sb.AppendLine($"Среднее число обслуживаемых заявок.\nLобс = ρ * Q = {rho:F3} * {Q:F3} = {L_obs:F3} ед.");
                        sb.AppendLine($"Число заявок, получивших отказ: λ * p_отк = {(lambda * p_otk):F2} заявок/мин.");
                        sb.AppendLine($"Номинальная производительность СМО: μ = {mu:F2} заявок/мин.");
                        sb.AppendLine($"Фактическая производительность СМО: {A:F3} / {mu:F2} = {(A / mu) * 100:F0}% от номинальной.");
                        break;

                    case 1: // С ограниченной очередью
                        if (string.IsNullOrWhiteSpace(LenghtQ.Text)) { MessageBox.Show("Введите длину очереди m."); return; }
                        int m1 = int.Parse(LenghtQ.Text);
                        sb.AppendLine($"\nИсчисляем показатели для одноканальной СМО с ограниченной очередью (m={m1}):");
                        sb.AppendLine($"Интенсивность нагрузки ρ = {rho:F3}");

                        double p0_1;
                        if (Math.Abs(rho - 1.0) < 0.0001)
                        {
                            p0_1 = 1.0 / (m1 + 2);
                            sb.AppendLine($"Вероятность того, что канал свободен (ρ=1): p0 = 1 / (m + 2) = {p0_1:F3}");
                        }
                        else
                        {
                            p0_1 = (1 - rho) / (1 - Math.Pow(rho, m1 + 2));
                            sb.AppendLine($"Вероятность того, что канал свободен: p0 = (1 - ρ) / (1 - ρ^(m+2)) = {p0_1:F3}");
                        }

                        double p_otk_1 = p0_1 * Math.Pow(rho, m1 + 1);
                        sb.AppendLine($"Вероятность отказа: P_отк = p0 * ρ^(m+1) = {p_otk_1:F3}");

                        double Q_1 = 1 - p_otk_1;
                        sb.AppendLine($"Относительная пропускная способность: Q = 1 - P_отк = {Q_1:F3}");

                        double A_1 = lambda * Q_1;
                        sb.AppendLine($"Абсолютная пропускная способность: A = λ * Q = {A_1:F3} заявок/мин.");

                        double L_och_1 = (rho * rho * (1 - (m1 + 1) * Math.Pow(rho, m1) + m1 * Math.Pow(rho, m1 + 1))) / ((1 - Math.Pow(rho, m1 + 2)) * (1 - rho));
                        sb.AppendLine($"Среднее число заявок в очереди: L_оч = {L_och_1:F3}");

                        double L_sist_1 = L_och_1 + Q_1;
                        sb.AppendLine($"Среднее число заявок в системе: L_сист = L_оч + Q = {L_sist_1:F3}");

                        double W_och_1 = L_och_1 / lambda;
                        sb.AppendLine($"Среднее время ожидания в очереди: T_оч = L_оч / λ = {W_och_1:F3} мин.");

                        double W_sist_1 = L_sist_1 / lambda;
                        sb.AppendLine($"Среднее время пребывания в системе: T_сист = L_сист / λ = {W_sist_1:F3} мин.");
                        break;

                    case 2: // С неограниченной очередью
                        sb.AppendLine("\nИсчисляем показатели для одноканальной СМО с неограниченной очередью:");
                        if (rho >= 1)
                        {
                            sb.AppendLine($"ВНИМАНИЕ: Интенсивность нагрузки ρ = {rho:F3} >= 1.");
                            sb.AppendLine("Система не справляется с потоком заявок, очередь будет расти до бесконечности.");
                            break;
                        }

                        double p0_2 = 1 - rho;
                        sb.AppendLine($"Вероятность того, что канал свободен: p0 = 1 - ρ = {p0_2:F3}");
                        sb.AppendLine("Вероятность отказа: P_отк = 0 (система без отказов)");
                        sb.AppendLine("Относительная пропускная способность: Q = 1");
                        sb.AppendLine($"Абсолютная пропускная способность: A = λ = {lambda:F3} заявок/мин.");

                        double L_och_2 = (rho * rho) / (1 - rho);
                        sb.AppendLine($"Среднее число заявок в очереди: L_оч = ρ^2 / (1 - ρ) = {L_och_2:F3}");

                        double L_sist_2 = rho / (1 - rho);
                        sb.AppendLine($"Среднее число заявок в системе: L_сист = ρ / (1 - ρ) = {L_sist_2:F3}");

                        double W_och_2 = L_och_2 / lambda;
                        sb.AppendLine($"Среднее время ожидания в очереди: T_оч = L_оч / λ = {W_och_2:F3} мин.");

                        double W_sist_2 = L_sist_2 / lambda;
                        sb.AppendLine($"Среднее время пребывания в системе: T_сист = L_сист / λ = {W_sist_2:F3} мин.");

                        if (!string.IsNullOrWhiteSpace(CountQ.Text))
                        {
                            int m2 = int.Parse(CountQ.Text);
                            double P_m = (1 - rho) * Math.Pow(rho, m2);
                            sb.AppendLine($"Вероятность нахождения m={m2} заявок в системе: P_m = (1-ρ)*ρ^m = {P_m:F4}");
                        }
                        break;
                }

                TextAnswer.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при расчете: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void Exsel_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TextAnswer.Text))
            {
                MessageBox.Show("Сначала выполните расчет (нажмите 'Решить').", "Внимание", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Add(Type.Missing);
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "Результаты СМО";
                string[] lines = TextAnswer.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < lines.Length; i++)
                {
                    worksheet.Cells[i + 1, 1] = lines[i];
                }
                worksheet.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при экспорте в Excel: {ex.Message}\nУбедитесь, что Excel установлен на компьютере.", "Ошибка экспорта", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }
        private void ChooseTimeIPZ_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
        }
        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            TextAnswer.Text = null;
            CountQ.Text = null;
            LenghtQ.Text = null;
            IPO.Text = null;
            IPZ.Text = null;
        }
    }
}