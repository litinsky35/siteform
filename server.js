const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');

const app = express();
const port = process.env.PORT || 3000;

// --- Добавлены парсеры для формы и JSON ---
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
// -----------------------------------

// Массивы для хранения данных
let fioList = [];          // Список фамилий
let registrations = [];    // Список записавшихся

// Настройка multer для загрузки файлов
const upload = multer({ dest: 'uploads/' });

// Статический маршрут для раздачи файлов (например, admin.html)
app.use(express.static(path.join(__dirname)));

// Обработчик для загрузки файла с фамилиями
app.post('/upload-fio', upload.single('file'), (req, res) => {
    const file = req.file;

    if (!file) {
        return res.status(400).json({ message: 'Нет файла для загрузки' });
    }

    const extname = path.extname(file.originalname).toLowerCase();
    let newFioList = [];

    try {
        if (extname === '.csv') {
            // Чтение CSV файла
            const csv = require('csv-parser');
            const fs = require('fs');
            fs.createReadStream(file.path)
                .pipe(csv())
                .on('data', (row) => {
                    newFioList.push(row.FIO); // Предполагаем, что в CSV есть колонка FIO
                })
                .on('end', () => {
                    fioList = newFioList; // Обновляем список фамилий
                    console.log('Загруженные фамилии (CSV):', fioList);
                    res.redirect('/admin'); // Перенаправляем на админ-панель после загрузки
                });
        } else if (extname === '.xlsx' || extname === '.xls') {
            // Чтение Excel файла
            const wb = xlsx.readFile(file.path);
            const ws = wb.Sheets[wb.SheetNames[0]];
            const data = xlsx.utils.sheet_to_json(ws);

            data.forEach(row => {
                if (row.FIO) {
                    newFioList.push(row.FIO); // Получаем фамилии из объекта
                }
            });
            fioList = newFioList; // Обновляем список фамилий
            console.log('Загруженные фамилии (Excel):', fioList);
            res.redirect('/admin'); // Перенаправляем на админ-панель после загрузки
        } else {
            res.status(400).json({ message: 'Неверный формат файла. Пожалуйста, загрузите CSV или Excel.' });
        }
    } catch (err) {
        res.status(500).json({ message: 'Ошибка при загрузке файла' });
    }
});

// Отдаем список фамилий для фронтенда
app.get('/get-fio-list', (req, res) => {
    res.json(fioList); // Отправляем список фамилий в ответ
});

// Обработчик для регистрации (записи на тестирование)
app.post('/register', (req, res) => {
    const { fio, date, time } = req.body;

    console.log('Получены данные от клиента:', fio, date, time);

    // Проверка, что все поля заполнены
    if (!fio || !date || !time) {
        return res.status(400).json({ message: 'Пожалуйста, заполните все поля.' });
    }

    // Проверяем, записался ли уже этот человек
    const alreadyRegistered = registrations.some(reg => reg.fio === fio);
    if (alreadyRegistered) {
        return res.status(400).json({ message: 'Вы уже записались на тестирование!' });
    }

    // Добавляем нового пользователя в список зарегистрированных
    registrations.push({ fio, date, time });
    console.log('Список записавшихся:', registrations); // Логируем список записавшихся

    // Ответ на успешную запись и редирект
    res.redirect('/success');  // Перенаправляем на страницу подтверждения
});

// Страница подтверждения успешной записи
app.get('/success', (req, res) => {
    res.sendFile(path.join(__dirname, 'success.html'));  // Путь к новому файлу success.html
});

// Обработка маршрута для админ-панели
app.get('/admin', (req, res) => {
    res.sendFile(path.join(__dirname, 'admin.html')); //Путь к файлу admin.html
});

// Обработчик для генерации Excel файла с записями
app.get('/download-registrations', (req, res) => {
    // Если нет записей, отправим сообщение
    if (registrations.length === 0) {
        return res.status(400).json({ message: 'Нет записей для скачивания.' });
    }

    // Создаём рабочий лист
    const ws = xlsx.utils.json_to_sheet(registrations);

    // Создаём рабочую книгу
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Записи');

    // Генерация Excel файла
    const filePath = path.join(__dirname, 'registrations.xlsx');
    xlsx.writeFile(wb, filePath);

    // Отправляем файл на скачивание
    res.download(filePath, 'registrations.xlsx', (err) => {
        if (err) {
            console.error('Ошибка при скачивании файла:', err);
        }

        // После скачивания, удалим файл с сервера
        fs.unlink(filePath, (unlinkErr) => {
            if (unlinkErr) {
                console.error('Ошибка при удалении файла:', unlinkErr);
            }
        });
    });
});

// Запуск сервера
app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
});