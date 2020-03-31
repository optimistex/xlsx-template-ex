import * as fs from "fs";
import { xlsxBuildByTemplate } from "./index";

const dummyData = {
  rootStore: {
    dummyArray: [
      {
        d1: 0,
        d2: 2
      },
      {
        d1: 3,
        d2: 4
      },
      {
        d1: 5,
        d2: 6
      }
    ],
    dummyString: "Im Dummy!",
    dummyObject: {
      veryDummyString: "Im sooo dummy!"
    }
  },
  rootValue: "root"
};

const testData = {
  reportBuildDate: "2018-05-18T18:11:14.227Z",
  taskCode: "ISE4",
  taskTechnicFio: "Полиненко Сергей специалист",
  taskName: "Инспекция №4",
  taskDescription: "CHECK LIST ОФИСНО-ТОРГОВАЯ, ПРОИЗВОДСТВЕННО-СКЛАСДКАЯ НЕДВИЖИМОСТЬ (КЛАССА НИЖЕ А И В)",
  taskDateStart: "15.05.2018",
  taskDateEnd: "25.05.2018",
  taskDateComplete: "16.05.2018",
  objectCode: "SE-1",
  objectName: "Офисное здание, 4 этажа",
  results: [
    {
      text: "Инженерные коммуникации: Электричество. Уточнить размер выделенной мощности (кВт).",
      answerText: "Нет электричества",
      comment: "",
      media: undefined,
      measureValue: "28"
    },
    {
      text: "Укажите количество мест для парковки и выберите тип",
      answerText: "Бесплатная наземная",
      comment: "",
      measureValue: "3"
    },
    {
      text:
        "Проверьте наличие на участке других объектов недвижимости. Если они есть, зафиксируйте длину и ширину каждого. Используйте знак препинания в качестве разделителя.",
      answerText: "Другие объекты отсутствуют",
      comment: "",
      measureValue: "34"
    },
    {
      text: "Укажите этажность здания",
      answerText: "более 2-х этажей",
      comment: "",
      measureValue: "5"
    },
    {
      text: "Действительный  адрес объекта соответствует адресу, указанному в документах?",
      answerText: "Производственно-складская",
      comment: "",
      media: null,
      measureValue: null
    },
    {
      text: "1 Конструктив здания: сфотографируйте фундамент и дефекты, если есть",
      answerText: "Нет дефектов",
      comment: "",
      measureValue: null,
      media: [
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461829069.jpeg",
          created: new Date(),
          gpsPos: "+++++"
        },
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461829069.jpeg",
          created: null,
          gpsPos: ""
        }
      ]
    },
    {
      text: "2 Конструктив здания: сфотографируйте стены и дефекты в них, если есть",
      answerText: "Нет дефектов в стенах",
      comment: "",
      measureValue: null,
      media: [
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461833615-1.jpeg",
          created: null,
          gpsPos: ""
        },
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461833615-2.jpeg",
          created: null,
          gpsPos: ""
        },
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461833615-3.jpeg",
          created: null,
          gpsPos: ""
        },
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461833615-4.jpeg",
          created: null,
          gpsPos: ""
        },
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461833615-5.jpeg",
          created: null,
          gpsPos: ""
        },
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461833615-6.jpeg",
          created: null,
          gpsPos: ""
        },
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461833615-7.jpeg",
          created: null,
          gpsPos: ""
        },
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461833615-8.jpeg",
          created: null,
          gpsPos: ""
        },
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461833615-9.jpeg",
          created: null,
          gpsPos: ""
        },
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461833615-10.jpeg",
          created: null,
          gpsPos: ""
        },
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461833615-11.jpeg",
          created: null,
          gpsPos: ""
        }
      ]
    },
    {
      text: "3 Конструктив здания: сфотографируйте перекрытия и дефекты в них, если есть",
      answerText: "Нет дефектов",
      comment: "",
      measureValue: null,
      media: [
        {
          fileName: "/data/www/itorum-backend/upload/answers/image-1526461838603.jpeg",
          created: null,
          gpsPos: ""
        }
      ]
    },
    {
      text: "Есть ли перепланировки? Уточните какие.",
      answerText: "Изменение материала внешних стен",
      comment: "",
      measureValue: null
    },
    {
      text: "Инженерные коммуникации: уточните тип отопления",
      answerText: "Централизованное",
      comment: "",
      measureValue: null
    },
    {
      text: "Инженерные коммуникации: уточните тип канализации",
      answerText: "Централизованное",
      comment: "",
      measureValue: null
    },
    {
      text: "Инженерные коммуникации: укажите тип газоснабжения",
      answerText: "Другое",
      comment: "",
      measureValue: null
    },
    {
      text: "Инженерные коммуникации: укажите тип водоснабжения",
      answerText: "Централизованное",
      comment: "",
      measureValue: null
    },
    {
      text: "Инженерные коммуникации: уточните наличие вентиляции",
      answerText: "Есть",
      comment: "",
      measureValue: null
    },
    {
      text: "Укажите какой подъездной путь",
      answerText: "Да",
      comment: "",
      measureValue: null
    },
    {
      text: "Инженерные коммуникации: уточните наличие системы кондиционирования и вентиляции",
      answerText: "Есть",
      comment: "",
      measureValue: null
    },
    {
      text: "Инженерные коммуникации: уточните наличие системы охраны и видеонаблюдения",
      answerText: "Есть",
      comment: "",
      measureValue: null
    },
    {
      text: "Инженерные коммуникации: уточните наличие системы пожарной сигнализации",
      answerText: "Есть",
      comment: "",
      measureValue: null
    },
    {
      text: "Инженерные коммуникации: другое (при наличии)",
      answerText: "Нет",
      comment: "",
      measureValue: null
    },
    {
      text: "Наличие трудно демонтируемого оборудования на объекте осмотра",
      answerText: "Нет",
      comment: "",
      measureValue: null
    },
    {
      text: "Линия домов",
      answerText: "Да",
      comment: "",
      measureValue: null
    },
    {
      text: "Окружающая застройка",
      answerText: "Да",
      comment: "",
      measureValue: null
    },
    {
      text: "Наличие отдельного входа",
      answerText: "Есть",
      comment: "",
      measureValue: null
    },
    {
      text: "Проведите осмотр состояния объекта",
      answerText: "Отличное",
      comment: "",
      measureValue: null
    },
    {
      text: "Доступ был предоставлен в 100% площадей объекта? Рассчитайте % площадей от общей площади объекта, в которые попасть не удалось.",
      answerText: "Да (доступ был к 100% площадей)",
      comment: "",
      measureValue: "100"
    },
    {
      text: "Укажите % текущего использования помещений объекта в соответствии с результатами визуального осмотра (в т.ч. основных арендаторов).",
      answerText: "Выполнено",
      comment: "",
      measureValue: "28"
    },
    {
      text: "Наличие вспомогательного оборудования на объекте осмотра",
      answerText: "Другое технологическое оборудование",
      comment: "",
      measureValue: null
    }
  ],
  steps: []
};

const testPipe = value => value + " test pipe";
const testPipeParams = (value, p1, p2) => value + " " + p1 + " " + p2;

xlsxBuildByTemplate(testData, __dirname + "/../test-data/xlsx-template-ex-0.xlsx")
  .then(buffer => {
    fs.writeFileSync("./out-0.xlsx", buffer);
  })
  .catch(error => {
    console.log("xlsxHelper error:", error);
  });

xlsxBuildByTemplate(testData, __dirname + "/../test-data/xlsx-template-ex-1.xlsx")
  .then(buffer => {
    fs.writeFileSync("./out-1.xlsx", buffer);
  })
  .catch(error => {
    console.log("xlsxHelper error:", error);
  });

xlsxBuildByTemplate(testData, __dirname + "/../test-data/xlsx-template-ex-2.xlsx")
  .then(buffer => {
    fs.writeFileSync("./out-2.xlsx", buffer);
  })
  .catch(error => {
    console.log("xlsxHelper error:", error);
  });

xlsxBuildByTemplate(dummyData, __dirname + "/../test-data/xlsx-template-ex-3.xlsx", { testPipe, testPipeParams })
  .then(buffer => {
    fs.writeFileSync("./out-3.xlsx", buffer);
  })
  .catch(error => {
    console.log("xlsxHelper error:", error);
  });
