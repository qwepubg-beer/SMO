using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
namespace WpfApp1.Pages
{
    /// <summary>
    /// Логика взаимодействия для Page2.xaml
    /// </summary>
    public partial class Page2 : Page
    {
        public Page2()
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
        private double Factorial(int n)
        {
            if (n == 0 || n == 1) return 1;
            double result = 1;
            for (int i = 2; i <= n; i++) result *= i;
            return result;
        }
        private void Answer_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(CountThread.Text) || string.IsNullOrWhiteSpace(IPZ.Text) || string.IsNullOrWhiteSpace(IPO.Text))
                {
                    MessageBox.Show("Заполните все обязательные поля (n, λ, μ / t_обс).", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int n = int.Parse(CountThread.Text);
                if (n <= 0) { MessageBox.Show("Количество каналов должно быть больше 0."); return; }

                // 1. Обработка Lambda (λ) - всегда интенсивность
                double lambda = double.Parse(IPZ.Text);
                if (ChooseTimeIPZ.SelectedIndex == 1) // Если для λ выбран "час"
                {
                    lambda = lambda / 60.0; // переводим в минуты
                }
                else if(ChooseTimeIPZ.SelectedIndex == 2)
                {
                    lambda = lambda / 1440;
                }
                double inputValue = double.Parse(IPO.Text);
                double mu, t_obs;

                if (ServiceType.SelectedIndex == 0) // Пользователь ввел μ (интенсивность)
                {
                    mu = inputValue;
                    if (Tops.SelectedIndex == 1) // Если выбраны часы для μ
                    {
                        mu = mu / 60.0; // переводим в заявки/мин
                    }
                    else if (Tops.SelectedIndex == 2)
                    {
                        mu = mu / 1440.0;
                    }
                    t_obs = 1.0 / mu;
                }
                else 
                {
                    t_obs = inputValue;
                    if (Tops.SelectedIndex == 1) // Если выбраны часы для t_обс
                    {
                        t_obs = t_obs * 60.0; // переводим в минуты
                    }
                    else if(Tops.SelectedIndex == 2)
                    {
                        t_obs = t_obs * 1440.0;
                    }
                    mu = 1.0 / t_obs;
                }

                // 3. Расчет нагрузок
                double r = lambda * t_obs; // Интенсивность нагрузки (равносильно lambda / mu)
                double rho = r / n;      

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Системы массового обслуживания (Многоканальные).");
                sb.AppendLine($"Количество каналов: n = {n}");
                sb.AppendLine($"Интенсивность потока заявок: λ = {lambda:F3} заявок/мин.");
                sb.AppendLine($"Интенсивность обслуживания: μ = {mu:F3} заявок/мин.");
                sb.AppendLine($"Среднее время обслуживания: t_обс = {t_obs:F3} мин.");
                sb.AppendLine($"Интенсивность нагрузки: r = λ * t_обс = {r:F3}");
                sb.AppendLine($"Коэффициент загрузки: ρ = r/n = {rho:F3}\n");

                switch (Choose.SelectedIndex)
                {
                    case 0:
                        sb.AppendLine("--- Многоканальная СМО с отказами ---");

                        double sumP0 = 0;
                        for (int k = 0; k <= n; k++)
                        {
                            sumP0 += Math.Pow(r, k) / Factorial(k);
                        }
                        double P0 = 1.0 / sumP0;
                        sb.AppendLine($"Вероятность того, что все каналы свободны: P0 = {P0:F4}");

                        double P_otk = (Math.Pow(r, n) / Factorial(n)) * P0;
                        sb.AppendLine($"Вероятность отказа (все каналы заняты): P_отк = {P_otk:F4} ({(P_otk * 100):F2}%)");

                        double Q = 1 - P_otk;
                        sb.AppendLine($"Относительная пропускная способность: Q = {Q:F4}");

                        double A = lambda * Q;
                        sb.AppendLine($"Абсолютная пропускная способность: A = {A:F3} заявок/мин.");

                        double k_busy = r * Q;
                        sb.AppendLine($"Среднее число занятых каналов: k = r * Q = {k_busy:F3}");

                        double t_sys = k_busy / lambda;
                        sb.AppendLine($"Среднее время пребывания заявки в системе: t_сист = {t_sys:F3} мин.");
                        break;

                    case 1:
                        if (string.IsNullOrWhiteSpace(LenghtQ.Text)) { MessageBox.Show("Введите длину очереди m."); return; }
                        int m = int.Parse(LenghtQ.Text);

                        sb.AppendLine($"--- Многоканальная СМО с ограниченной очередью (m = {m}) ---");

                        double sumBase = 0;
                        for (int k = 0; k <= n; k++) sumBase += Math.Pow(r, k) / Factorial(k);

                        double sumQueue = 0;
                        for (int k = 1; k <= m; k++) sumQueue += Math.Pow(r / n, k);

                        double P0_lim = 1.0 / (sumBase + (Math.Pow(r, n) / Factorial(n)) * sumQueue);
                        sb.AppendLine($"Вероятность того, что все каналы свободны: P0 = {P0_lim:F4}");

                        double P_otk_lim = (Math.Pow(r, n + m) / (Factorial(n) * Math.Pow(n, m))) * P0_lim;
                        sb.AppendLine($"Вероятность отказа: P_отк = {P_otk_lim:F4} ({(P_otk_lim * 100):F2}%)");

                        double Q_lim = 1 - P_otk_lim;
                        sb.AppendLine($"Относительная пропускная способность: Q = {Q_lim:F4}");

                        double A_lim = lambda * Q_lim;
                        sb.AppendLine($"Абсолютная пропускная способность: A = {A_lim:F3} заявок/мин.");

                        double L_och = 0;
                        double coeff = (P0_lim * Math.Pow(r, n)) / Factorial(n);
                        for (int k = 1; k <= m; k++) L_och += k * Math.Pow(r / n, k);
                        L_och *= coeff;
                        sb.AppendLine($"Средняя длина очереди: L_оч = {L_och:F3}");

                        double L_sys = L_och + r * Q_lim;
                        sb.AppendLine($"Среднее число заявок в системе: L_сист = {L_sys:F3}");

                        double W_och = L_och / lambda;
                        sb.AppendLine($"Среднее время ожидания в очереди: T_оч = {W_och:F3} мин.");

                        double W_sys = L_sys / lambda;
                        sb.AppendLine($"Среднее время в системе: T_сист = {W_sys:F3} мин.");
                        break;

                    case 2:
                        sb.AppendLine("--- Многоканальная СМО с неограниченной очередью ---");
                        if (rho >= 1)
                        {
                            sb.AppendLine($"ВНИМАНИЕ: Коэффициент загрузки ρ = {rho:F3} >= 1.");
                            sb.AppendLine("Система не справляется с потоком. Очередь будет бесконечно расти.");
                            break;
                        }

                        double sumBaseUn = 0;
                        for (int k = 0; k < n; k++) sumBaseUn += Math.Pow(r, k) / Factorial(k);

                        double P0_un = 1.0 / (sumBaseUn + (Math.Pow(r, n) / Factorial(n)) * (1.0 / (1 - rho)));
                        sb.AppendLine($"Вероятность того, что все каналы свободны: P0 = {P0_un:F4}");

                        sb.AppendLine("Вероятность отказа: P_отк = 0 (очередь не ограничена)");
                        sb.AppendLine("Относительная пропускная способность: Q = 1");
                        sb.AppendLine($"Абсолютная пропускная способность: A = λ = {lambda:F3} заявок/мин.");

                        double P_wait = (Math.Pow(r, n) / (Factorial(n) * (1 - rho))) * P0_un;
                        sb.AppendLine($"Вероятность ожидания в очереди: P_ож = {P_wait:F4}");

                        double L_och_un = P_wait * (rho / (1 - rho));
                        sb.AppendLine($"Средняя длина очереди: L_оч = {L_och_un:F3}");

                        double L_sys_un = L_och_un + r;
                        sb.AppendLine($"Среднее число заявок в системе: L_сист = {L_sys_un:F3}");

                        double W_och_un = L_och_un / lambda;
                        sb.AppendLine($"Среднее время ожидания в очереди: T_оч = {W_och_un:F3} мин.");

                        double W_sys_un = L_sys_un / lambda;
                        sb.AppendLine($"Среднее время в системе: T_сист = {W_sys_un:F3} мин.");

                        if (!string.IsNullOrWhiteSpace(CountQ.Text))
                        {
                            int m2 = int.Parse(CountQ.Text);
                            double P_m = P_wait * Math.Pow(rho, m2) * (1 - rho);
                            sb.AppendLine($"Вероятность того, что в системе ровно {n + m2} заявок (n каналов + {m2} в очереди): {P_m:F4}");
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
                worksheet.Name = "Многоканальная СМО";

                string[] lines = TextAnswer.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                for (int i = 0; i < lines.Length; i++)
                {
                    worksheet.Cells[i + 1, 1] = lines[i];
                }

                worksheet.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта в Excel: {ex.Message}\nУбедитесь, что Excel установлен.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            TextAnswer.Text = null;
            CountQ.Text=null;
            LenghtQ.Text=null;
            IPO.Text=null;
            IPZ.Text=null;
            CountThread.Text=null;
        }
    }
}
