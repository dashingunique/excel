# excel
Excel import/export for Php

Install via composer:

```
composer require dashingunique/excel
```
## 导出文件信息

将文件导出到 `.csv`(.xlsx .obs) 文件:

```php
use dashingunique\excel\DashingExcel;
use app\model\User;

// Load users
$users = new User()->select();
$users = uniqueCollection($users);
// Export all users
(new DashingExcel($users))->export('file.csv');
```


仅导入指定信息的列
```php
$users = new User()->select();
$users = uniqueCollection($users);
(new DashingExcel($users))->export('users.csv', function ($user) {
    return [
        'Email' => $user['email'],
        'First Name' => $user['firstname'],
        'Last Name' => strtotime($user['create_time']),
    ];
});
```

## 导入文件信息

导入文件信息
```php
$collection = (new DashingExcel())->configureCsv(';', '#', '\n', 'gbk')->import('file.csv');
```

导入文件并写入数据库
```php
$users = (new DashingExcel())->import('file.xlsx', function ($line) {
    return (new User())->create([
        'name' => $line['Name'],
        'email' => $line['Email']
    ]);
});
```

