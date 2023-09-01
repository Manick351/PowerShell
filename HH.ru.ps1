<#
This script is written in powershell and is bound to the city id, it is also a test version for more flexible filters and settings.
Initially, the script opens an empty window, while it has already collected data, it remains to filter them according to your needs.
A limited number of tabs is assumed so far, if necessary, they can be increased or decreased.
Autor manick351@live.com

該腳本是用powershell編寫的，並綁定了城市id，它也是更靈活的過濾器和設置的測試版本。
最初，該腳本會打開一個空窗口，雖然它已經收集了數據，但仍會根據您的需要過濾它們。
到目前為止，假設選項卡數量有限，如有必要，可以增加或減少它們。
作者 manick351@live.com

Этот скрипт написан на powershell и привязан к идентификатору города, это также тестовая версия для более гибких фильтров и настроек.
Изначально скрипт открывает пустое окно, пока он уже собрал данные, осталось отфильтровать их по вашим потребностям.
Пока предполагается ограниченное количество вкладок, при необходимости их можно увеличить или уменьшить.
Автор manick351@live.com

#>

Add-Type -AssemblyName System.Windows.Forms

$form = New-Object Windows.Forms.Form
$form.Text = "HH.ru Vacancies"
$form.Width = 1400
$form.Height = 900

$dataGridView = New-Object Windows.Forms.DataGridView
$dataGridView.AutoGenerateColumns = $false

$dataGridView.Location = New-Object System.Drawing.Point(10, 10)
$dataGridView.Size = New-Object System.Drawing.Size(1350, 650)

# Добавление столбца с ссылками на URL
$urlColumn = New-Object Windows.Forms.DataGridViewLinkColumn
$urlColumn.HeaderText = "URL"
$urlColumn.DataPropertyName = "Url"
$urlColumn.UseColumnTextForLinkValue = $false
$urlColumn.Width = 100
$dataGridView.Columns.Add($urlColumn) | Out-Null

# Добавление обработчика события для открытия URL в браузере
$dataGridView.add_CellContentClick({
    param (
        [System.Object]$sender,
        [System.Windows.Forms.DataGridViewCellEventArgs]$e
    )

    # Проверка, что клик был в столбце URL
    if ($e.ColumnIndex -eq $urlColumn.Index) {
        $url = $dataGridView.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value.ToString()
        Start-Process $url
    }
})

$idColumn = New-Object Windows.Forms.DataGridViewTextBoxColumn
$idColumn.HeaderText = "ID"
$idColumn.DataPropertyName = "Id"
$idColumn.Width = 60
$dataGridView.Columns.Add($idColumn) | Out-Null

$salaryColumn = New-Object Windows.Forms.DataGridViewTextBoxColumn
$salaryColumn.HeaderText = "Зарплата"
$salaryColumn.DataPropertyName = "Salary"
$salaryColumn.Width = 120
$dataGridView.Columns.Add($salaryColumn) | Out-Null

$nameColumn = New-Object Windows.Forms.DataGridViewTextBoxColumn
$nameColumn.HeaderText = "Название"
$nameColumn.DataPropertyName = "Name"
$nameColumn.Width = 370
$dataGridView.Columns.Add($nameColumn) | Out-Null

$createdAtColumn = New-Object Windows.Forms.DataGridViewTextBoxColumn
$createdAtColumn.HeaderText = "Дата публикации"
$createdAtColumn.DataPropertyName = "CreatedAt"
$createdAtColumn.Width = 110
$dataGridView.Columns.Add($createdAtColumn) | Out-Null

$statusColumn = New-Object Windows.Forms.DataGridViewTextBoxColumn
$statusColumn.HeaderText = "Статус"
$statusColumn.DataPropertyName = "Status"
$statusColumn.Width = 55
$dataGridView.Columns.Add($statusColumn) | Out-Null

$employerColumn = New-Object Windows.Forms.DataGridViewTextBoxColumn
$employerColumn.HeaderText = "Компания"
$employerColumn.DataPropertyName = "Employer"
$employerColumn.Width = 320
$dataGridView.Columns.Add($employerColumn) | Out-Null

$cityColumn = New-Object Windows.Forms.DataGridViewTextBoxColumn
$cityColumn.HeaderText = "Город"
$cityColumn.DataPropertyName = "City"
$cityColumn.Width = 180
$dataGridView.Columns.Add($cityColumn) | Out-Null

$form.Controls.Add($dataGridView)

# Добавление элемента управления TextBox для фильтрации
$filterTextBox = New-Object Windows.Forms.TextBox
$filterTextBox.Location = [System.Drawing.Point]::new(10, 715)
$filterTextBox.Width = 200
$form.Controls.Add($filterTextBox)

# Добавление кнопки "Фильтр"
$filterButton = New-Object Windows.Forms.Button
$filterButton.Text = "Фильтр"
$filterButton.Location = New-Object System.Drawing.Point(220, 715)
$filterButton.Add_Click({
    $search_value = $filterTextBox.Text.ToLower()
    foreach ($row in $dataGridView.Rows) {
        if ($row.Index -eq $dataGridView.NewRowIndex) {
            continue
        }
        $isVisible = $false
        foreach ($cell in $row.Cells) {
            if ($cell.Value.ToString().ToLower() -like "*$search_value*") {
                $isVisible = $true
                break
            }
        }
        $row.Visible = $isVisible
    }
})

$form.Controls.Add($filterButton)

# Добавление кнопки "OK"
$okButton = New-Object Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(620, 700)
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($okButton)

# Добавление кнопки "Cancel"
$cancelButton = New-Object Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Location = New-Object System.Drawing.Point(760, 700)
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$cancelButton.Add_Click({ $form.Close() })
$form.Controls.Add($cancelButton)

# URL API hh.ru
$url = "https://api.hh.ru/vacancies"
# Заголовки запроса
$headers = @{
    'User-Agent' = 'api-test-agent'
}
# Количество результатов на странице
$per_page = 100
# URL API hh.ru для поиска компаний
$employersUrl = "https://api.hh.ru/employers"

# Параметры запроса
$requestParams = @{
    area = "24"  # Идентификатор региона 24
    only_with_vacancies = $true
    per_page = 10  # Устанавливаем количество записей на странице
}

# Создание UriBuilder и добавление параметров
$uriBuilder = New-Object System.UriBuilder($employersUrl)
#$queryParams = [System.Web.HttpUtility]::ParseQueryString("")  # Создание объекта для хранения параметров
$queryParams = @{}  # Создание объекта для хранения параметров

foreach ($param in $requestParams.GetEnumerator()) {
    $queryParams[$param.Key] = $param.Value
}

$uriBuilder.Query = $queryParams.ToString()

# Формирование полной ссылки с параметрами
$fullUrl = $uriBuilder.ToString()

# Выполнение запроса для получения списка компаний с вакансиями
$employersResponse = Invoke-RestMethod -Uri $fullUrl -Headers $headers

# Определение количества записей и записей на странице
$totalRecords = [Math]::Min($employersResponse.found, 5000)  # Ограничение до 5000 компаний
$recordsPerPage = 10  # Установите желаемое количество записей на странице
$pages = [Math]::Ceiling($totalRecords / $recordsPerPage)

# Создание списка компаний
$employersList = @()

# Обработка всех страниц с компаниями
for ($page = 0; $page -lt $pages; $page++) {
    $queryParams["page"] = $page
    $queryParams["per_page"] = $recordsPerPage

    $clonedUriBuilder = [System.UriBuilder]::new($employersUrl)
    $query = [System.Web.HttpUtility]::ParseQueryString("")
    foreach ($param in $queryParams.GetEnumerator()) {
        $query[$param.Key] = $param.Value
    }
    $clonedUriBuilder.Query = $query.ToString()
    $fullUrl = $clonedUriBuilder.Uri.ToString()

    try {
        $employersResponse = Invoke-RestMethod -Uri $fullUrl -Headers $headers

        $employersList += $employersResponse.items
    } catch {
        Write-Host "An error occurred while fetching companies: $_"
    }
}

# Создание ComboBox для выбора компании
$companyComboBox = New-Object Windows.Forms.ComboBox
$companyComboBox.Location = New-Object System.Drawing.Point(10, 760)
$companyComboBox.Width = 400
$form.Controls.Add($companyComboBox)

# Заполнение ComboBox списком компаний
foreach ($employer in $employersList) {
    $companyComboBox.Items.Add($employer.name) | Out-Null
}

# Создание кнопки "Поиск вакансий"
$searchVacanciesButton = New-Object Windows.Forms.Button
$searchVacanciesButton.Text = "Поиск вакансий"
$searchVacanciesButton.Location = New-Object System.Drawing.Point(420, 760)
$searchVacanciesButton.Size = New-Object System.Drawing.Size(100, 25)
$form.Controls.Add($searchVacanciesButton)

$searchVacanciesButton.Add_Click({
    if ($companyComboBox.SelectedIndex -ge 0) {
        $selectedCompany = $companyComboBox.SelectedItem
        
        # Проверяем, что $selectedCompany не пустой
        if (-not [string]::IsNullOrEmpty($selectedCompany)) {
            $employer = $employersList | Where-Object { $_.name -eq $selectedCompany }
            
            # Проверяем, что $employer не пустой
            if ($employer) {
                $uri = [System.Uri]::new($employer.vacancies_url)
                $response = Invoke-RestMethod -Uri $uri -Headers $headers
                $results = @()
                foreach ($vacancy in $response.items) {
                    $status = if ($vacancy.type.id -eq "open") { "Открыта" } else { "Закрыта" }
                    $createdAt = [DateTime]::ParseExact($vacancy.created_at, "yyyy-MM-ddTHH:mm:sszzz", $null)
                        # Проверяем наличие информации о зарплате
        $salary = if ($vacancy.salary) {
            if ($vacancy.salary.from -and $vacancy.salary.to) {
                "$($vacancy.salary.from) - $($vacancy.salary.to)"
                } elseif ($vacancy.salary.from) {
                    "От $($vacancy.salary.from)"
                } elseif ($vacancy.salary.to) {
                    "До $($vacancy.salary.to)"
                } else {
                    "Не указана"
        }
    } else {
        "Не указана"
    }
	# Создание объекта данных и добавление в результаты
    $rowData = New-Object PSObject -Property @{
        Url = $vacancy.alternate_url
        Id = $vacancy.id
        Salary = $salary
        Name = $vacancy.name
        CreatedAt = $createdAt.ToString("yyyy-MM-dd HH:mm:ss")
        Status = $status
        Employer = $vacancy.employer.name
        City = $vacancy.area.name
    }
    $results += $rowData
}
                # Заполняем DataGridView данными
                $dataGridView.Rows.Clear()
                $results | ForEach-Object {
                    $dataGridView.Rows.Add($_.Url, $_.Id,  $_.Salary, $_.Name, $_.CreatedAt, $_.Status, $_.Employer, $_.City) | Out-Null
                }
            }
        }
    }
})

# Выполнение запроса для получения списка категорий и ролей
$response = Invoke-RestMethod -Uri "https://api.hh.ru/professional_roles" -Headers $headers

# Создание ComboBox для выбора категории и роли
$roleComboBox = New-Object Windows.Forms.ComboBox
$roleComboBox.Location = New-Object System.Drawing.Point(10, 675)
$roleComboBox.Width = 400
$form.Controls.Add($roleComboBox)

# Заполнение ComboBox списком категорий и ролей
foreach ($category in $response.categories) {
    foreach ($role in $category.roles) {
        $roleComboBox.Items.Add("$($category.name) - $($role.name)") | Out-Null
    }
}

# Добавление кнопки "Поиск"
$searchButton = New-Object Windows.Forms.Button
$searchButton.Text = "Поиск"
$searchButton.Location = New-Object System.Drawing.Point(420, 675)
$form.Controls.Add($searchButton)

$searchButton.Add_Click({
    if ($roleComboBox.SelectedIndex -ge 0) {
        $selectedRole = $roleComboBox.SelectedItem -split " - "
        $selectedCategory = $selectedRole[0]
        $selectedRoleName = $selectedRole[1]
        
        $category_id = $response.categories | Where-Object { $_.name -eq $selectedCategory } |
                       Select-Object -ExpandProperty id
        $role_id = $response.categories | Where-Object { $_.name -eq $selectedCategory } |
                    Select-Object -ExpandProperty roles | Where-Object { $_.name -eq $selectedRoleName } |
                   Select-Object -ExpandProperty id
        
        # Выполнение запроса для получения вакансий для выбранной категории и роли
        $results = @()

        for ($page = 0; $page -lt 5; $page++) {
            $uriBuilder = [System.UriBuilder] $url
            $uriBuilder.Query = "page=$page&per_page=$per_page&professional_area=$category_id&professional_role=$role_id"
            $uri = $uriBuilder.Uri

            $response = Invoke-RestMethod -Uri $uri -Headers $headers

foreach ($vacancy in $response.items) {
    $status = if ($vacancy.type.id -eq "open") { "Открыта" } else { "Закрыта" }
    $createdAt = [DateTime]::ParseExact($vacancy.created_at, "yyyy-MM-ddTHH:mm:sszzz", $null)

    # Проверяем наличие информации о зарплате
    $salary = if ($vacancy.salary) {
        if ($vacancy.salary.from -and $vacancy.salary.to) {
            "$($vacancy.salary.from) - $($vacancy.salary.to)"
        } elseif ($vacancy.salary.from) {
            "От $($vacancy.salary.from)"
        } elseif ($vacancy.salary.to) {
            "До $($vacancy.salary.to)"
        } else {
            "Не указана"
        }
    } else {
        "Не указана"
    }
	# Создание объекта данных и добавление в результаты
    $rowData = New-Object PSObject -Property @{
        Url = $vacancy.alternate_url
        Id = $vacancy.id
        Salary = $salary
        Name = $vacancy.name
        CreatedAt = $createdAt.ToString("yyyy-MM-dd HH:mm:ss")
        Status = $status
        Employer = $vacancy.employer.name
        City = $vacancy.area.name
    }
    $results += $rowData
}
        }

        # Заполнение DataGridView данными
        $dataGridView.Rows.Clear()
        $results | ForEach-Object {
            $dataGridView.Rows.Add($_.Url, $_.Id,  $_.Salary, $_.Name, $_.CreatedAt, $_.Status, $_.Employer, $_.City) | Out-Null
        }
    }
})

# Запуск формы и обработка результатов
$result = $form.ShowDialog()

# Завершение работы скрипта
$form.Dispose()
