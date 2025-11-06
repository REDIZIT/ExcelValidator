using System;
using System.Collections.Generic;
using System.Reflection;

public class TestContext
{
    public TestAttribute attr;
    public MethodInfo testMethod;
    public MetaTable t;
    public Result result;
    public Assert assert => result.assert;
    public Status status;

    private HashSet<string> websites = new()
    {
        "РАД", "Торги России", "ДОМ.РФ", "Фонд Имущества", "Балтийская электронная площадка", "другое"
    };
    private HashSet<string> gosServices = new()
    {
        "Петербургская недвижимость", "Росреестр", "ГУИОН"
    };

    private Column year;
    private Column src1, src2, src3;
    private Column doc, screens, deal;
    private Column com, jud, link, portal;
    private Column coef, price;

    private const float LOW_PRICE = 10;

    public enum Status
    {
        ReadyToRun,
        Running,
        Ok,
        Failed,
        Exception,
    }

    public TestContext(MetaTable t)
    {
        this.t = t;

        result = new(t);
        result.ctx = this;

        year = t.GetColumn("Год");
        src1 = t.GetColumn("Источник 1");
        src2 = t.GetColumn("Источник 2");
        src3 = t.GetColumn("Источник 3");
        doc = t.GetColumn("Докум.");
        screens = t.GetColumn("Кол-во основных скринов");
        deal = t.GetColumn("Вид сделки");
        com = t.GetColumn("Комиссия");
        jud = t.GetColumn("Суд");
        link = t.GetColumn("Ссылка");
        portal = t.GetColumn("Портал");
        coef = t.GetColumn("Коэф. превышения");
        price = t.GetColumn("Нач. общ. стоим. лота");
    }

    public void Run()
    {
        if (status != Status.ReadyToRun)
        {
            throw new($"Can not run same {nameof(TestContext)} twice. Create new context.");
        }

        status = Status.Running;

        try
        {
            testMethod.Invoke(this, null);
            status = result.IsGood ? Status.Ok : Status.Failed;
        }
        catch (Exception err)
        {
            status = Status.Exception;
            result.exception = err;
        }
    }

    private IEnumerable<int> EachRow()
    {
        for (int y = 0; y < t.rows; y++)
        {
            assert.ctx.y = y;
            yield return y;
        }
    }

    [Test("Год, коды, лоты (1.)", "Общие")]
    private void Check1()
    {
        Column[] columnsToCheck = new[]
        {
            t.GetColumn("Код лота"),
            t.GetColumn("Лот"),
            t.GetColumn("Код ОН"),
            t.GetColumn("Уник №"),
        };

        foreach (int y in EachRow())
        {
            string? yearString = t.At(year, y).ToString();

            if (string.IsNullOrWhiteSpace(yearString))
            {
                result.Mark(year, y, "'Год' не заполнен");
                continue;
            }

            for (int i = 0; i < columnsToCheck.Length; i++)
            {
                assert.That(columnsToCheck[i], s => s.Contains(yearString));
            }
        }
    }

    [Test("Прочее имущество (2.)", "Общие")]
    private void Check2()
    {
        Column filterColumn = t.GetColumn("Прочее имущество, руб.");

        foreach (int y in EachRow())
        {
            if (assert.IsNotEmpty(filterColumn))
            {
                string src1_value = t.At<string>(src1, y);

                if (src1_value.StartsWith("Торги "))
                {
                    assert.Equal(src2, y, "ГБУ КО");
                    assert.That(src3, y, websites.Contains);
                    assert.NotZeroOrEmpty(doc);
                    assert.NotZeroOrEmpty(screens);
                }
                else if (src1_value == "Рынок")
                {
                    assert.Equal(src2, y, "Росреестр");
                    assert.Empty(src3, y);
                    assert.Equal(deal, y, "Сделка");
                }
            }
        }
    }

    [Test("Нач. общ. стоим. лота (3.)", "Общие")]
    private void Check3()
    {
        foreach (int y in EachRow())
        {
            if (assert.IsNotEmpty(price))
            {
                assert.NotZeroOrEmpty(coef);
                assert.That(src1, s => s.StartsWith("Торги "));
                assert.Equal(src2, "ГБУ КО");
                assert.That(src3, websites.Contains);
                assert.NotZeroOrEmpty(doc);
                assert.NotZeroOrEmpty(screens);
            }
        }
    }

    [Test("Коэф. превышения (4.)", "Общие")]
    private void Check4()
    {
        foreach (int y in EachRow())
        {
            if (assert.IsNotEmpty(coef))
            {
                assert.NotZeroOrEmpty(price);
                assert.That(src1, s => s.StartsWith("Торги "));
                assert.Equal(src2, "ГБУ КО");
                assert.That(src3, s => websites.Contains(s));
                assert.NotZeroOrEmpty(doc);
                assert.NotZeroOrEmpty(screens);
            }
        }
    }

    [Test("Фактическое использование ЗУ (5.)", "Общие")]
    private void Check5()
    {
        var fact = t.GetColumn("Факт. использование ЗУ");

        foreach (int y in EachRow())
        {
            assert.NotEmpty(fact);
        }
    }

    [Test("Доля ЗУ (6.)", "Общие")]
    private void Check6()
    {
        var zu = t.GetColumn("Доля ЗУ");
        var type = t.GetColumn("Вид ОН");

        foreach (int y in EachRow())
        {
            if (assert.IsZeroOrEmpty(zu))
            {
                assert.Equal(type, "OKS");
            }
        }
    }

    [Test("Зона ПЗЗ и СЭР (7. и 8.)", "Общие")]
    private void Check7()
    {
        var zone = t.GetColumn("Зона ПЗЗ");
        var serDaSer = t.GetColumn("Код СЭР");

        foreach (int y in EachRow())
        {
            assert.That(zone, s => s != "нет" && string.IsNullOrWhiteSpace(s) == false);
            assert.That(serDaSer, s => s != "нет" && string.IsNullOrWhiteSpace(s) == false);
        }
    }

    [Test("Источник 1 (9.)", "Общие")]
    private void Check10()
    {
        var zu = t.GetColumn("Право ЗУ");
        var zuNote = t.GetColumn("Примечание (ЗУ)");

        foreach (int y in EachRow())
        {
            if (assert.IsEqual(src1, "База данных"))
            {
                assert.Equal(src2, "ГБУ КО");
                assert.Equal(src3, "Комиссия/Суд по оспариванию КС");
                assert.Equal(link, "---");
                assert.Equal(portal, "---");
                assert.NotZeroOrEmpty(doc);
                assert.NotZeroOrEmpty(screens);
                assert.Equal(deal, "Сделка");
                assert.NotContains(zu, "Аренда");
            }
            else if (assert.Is(src1, s => s.StartsWith("Торги ")))
            {
                assert.Equal(src2, "ГБУ КО");
                assert.That(src3, s => websites.Contains(s));
                assert.NotZeroOrEmpty(doc);
                assert.NotZeroOrEmpty(screens);
                assert.NotEmpty(price);
                assert.NotEmpty(coef);
            }
            else if (assert.IsEqual(src1, "Рынок") && assert.IsEqual(src2, "ГБУ КО"))
            {
                assert.NotEmpty(doc);

                assert.NotEmpty(src3);
                if (assert.Is(src3, s => websites.Contains(s)))
                {
                    assert.That(zuNote, s =>
                    {
                        string sl = s.ToLower();
                        return sl.Contains("продажа без торгов") || sl.Contains("прямая продажа");
                    });
                }

                assert.NotZeroOrEmpty(screens);
            }
            else if (assert.IsEqual(src1, "Рынок") && assert.IsEqual(src2, "Росреестр"))
            {
                assert.Empty(src3);
                assert.Contains(link, "rosreestr");
                assert.Equal(portal, "rosreestr.gov.ru");
                assert.Equal(deal, "Сделка");
            }
        }
    }

    [Test("Источник 2 (10.)", "Общие")]
    private void Check11()
    {
        var zuNote = t.GetColumn("Примечание (ЗУ)");

        foreach (int y in EachRow())
        {
            if (assert.IsEqual(src2, "Росреестр"))
            {
                assert.Equal(src1, "Рынок");
                assert.Empty(src3);
                assert.Contains(link, "rosreestr");
                assert.Equal(portal, "rosreestr.gov.ru");
                assert.Equal(deal, "Сделка");
                assert.ZeroOrEmpty(screens);
            }
            else if (assert.IsEqual(src2, "ГБУ КО"))
            {
                assert.NotEmpty(doc);
                assert.NotZeroOrEmpty(screens);

                if (assert.IsEqual(src1, "Рынок") && assert.Is(src3, s => websites.Contains(s)))
                {
                    assert.That(zuNote, s =>
                    {
                        string sl = s.ToLower();
                        return sl.Contains("продажа без торгов") || sl.Contains("прямая продажа");
                    });
                }
            }
        }
    }

    [Test("Докум. (11.)", "Общие")]
    private void Check12()
    {
        foreach (int y in EachRow())
        {
            if (assert.Is(src2, s => gosServices.Contains(s)))
            {
                // can be zero or empty
            }
            else
            {
                assert.NotZeroOrEmpty(doc);
            }
        }
    }

    [Test("Скрины (12.)", "Общие")]
    private void Check13()
    {
        foreach (int y in EachRow())
        {
            if (assert.Is(src2, s => gosServices.Contains(s)))
            {
                // can be zero or empty
            }
            else
            {
                assert.NotZeroOrEmpty(screens);
            }
        }
    }

    [Test("Снос (13.)", "Общие")]
    private void Check14()
    {
        var dem = t.GetColumn("Снос");
        var demPurpose = t.GetColumn("Назнач. дома под снос");
        var area = t.GetColumn("Площадь ОКС");
        var fact = t.GetColumn("Факт. использование ОКС");
        var type = t.GetColumn("Вид ОН");
        var uni = t.GetColumn("Единый лот");
        var add = t.GetColumn("Дополн. фильтр");

        foreach (int y in EachRow())
        {
            if (assert.IsEqual(dem, "Снос"))
            {
                assert.NotEmpty(demPurpose);
                assert.NotZeroOrEmpty(area);

                if (assert.IsZeroOrEmpty(area))
                {
                    assert.Contains(add, "Фиктивный КН");
                }
            }
            else if (assert.Is(dem, s => s == "Реконструкция" || s == "Ремонт"))
            {
                assert.Empty(demPurpose);

                if (assert.IsEqual(type, "ONS") || (assert.IsEqual(type, "ENC") && uni.Equals("Единый лот")))
                {
                    assert.ZeroOrEmpty(area);
                }
                else
                {
                    assert.NotZeroOrEmpty(area);
                }
            }
        }
    }

    [Test("Факт. исп. ОКС (14.)", "Общие")]
    private void FactOKS()
    {
        var fact = t.GetColumn("Факт. использование ОКС");
        var kad = t.GetColumn("Кад. номер ОКС");
        var add = t.GetColumn("Дополн. фильтр");

        foreach (int y in EachRow())
        {
            if (assert.IsNotEmpty(fact))
            {
                assert.NotEmpty(kad);
            }
            else
            {
                if (assert.IsEmpty(kad) || assert.IsContains(add, "Фиктивный КН"))
                {
                    // ok
                }
                else
                {
                    result.Mark(kad, y, $"Для пустого '{fact.name}' ожидалось, что '{kad.name}' = <пусто> или '{add.name}' будет содержать 'Фиктивный КН'");
                }
            }
        }
    }

    [Test("Площадь ОКС (15.)", "Общие")]
    private void AreaOKS()
    {
        var area = t.GetColumn("Площадь ОКС");
        var kad = t.GetColumn("Кад. номер ОКС");
        var add = t.GetColumn("Дополн. фильтр");

        foreach (int y in EachRow())
        {
            if (assert.IsEmpty(kad) || assert.IsContains(add, "Фиктивный КН"))
            {
                //assert.ZeroOrEmpty(area); // may be ZeroOrEmpty or must be ZeroOrEmpty?
            }
            else
            {
                assert.NotZeroOrEmpty(area);
            }
        }
    }

    [Test("Подсегмент (16.)", "Общие")]
    private void Subsegment()
    {
        var sub = t.GetColumn("Подсегмент");
        var type = t.GetColumn("Вид ОН");

        foreach (int y in EachRow())
        {
            if (assert.IsEqual(sub, "ОГОРОД"))
            {
                assert.Equal(type, "ZU");
            }
        }
    }


    [Test("ЗУ и Rnt (1.-6.)", "ЗУ")]
    private void ZuRnt()
    {
        var type = t.GetColumn("Вид ОН");
        var additionalFilter = t.GetColumn("Дополн. фильтр");
        var uni = t.GetColumn("Единый лот");

        var sellPrice = t.GetColumn("Цена продажи ЗУ");
        var unitPrice = t.GetColumn("Уд. цена на  1 кв.м ЗУ");
        var areaLot = t.GetColumn("Площадь лота");
        var areaZu = t.GetColumn("Площадь ЗУ");
        var fractZu = t.GetColumn("Доля ЗУ");
        var areaFractZu = t.GetColumn("Площадь доли ЗУ");

        foreach (int y in EachRow())
        {
            bool filter = assert.IsEqual(type, "ZU") || assert.IsEqual(type, "Rnt");
            if (!filter) continue;

            filter = assert.IsNotEqual(additionalFilter, "АРХИВ");
            if (!filter) continue;

            filter = assert.IsEmpty(uni);
            if (!filter) continue;

            assert.NotZeroOrEmpty(sellPrice);
            assert.NotZeroOrEmpty(unitPrice);
            assert.NotZeroOrEmpty(areaLot);
            assert.NotZeroOrEmpty(areaZu);
            assert.NotZeroOrEmpty(fractZu);
            assert.NotZeroOrEmpty(areaFractZu);
        }
    }

    [Test("ЗУ и Rnt (7.)", "ЗУ")]
    private void ZuRnt2()
    {
        var type = t.GetColumn("Вид ОН");
        var additionalFilter = t.GetColumn("Дополн. фильтр");
        var uni = t.GetColumn("Единый лот");

        var sellPriceLot = t.GetColumn("Цена продажи ОКС (лот)");
        var priceV = t.GetColumn("Стоимость ОКС указана отдельно");
        var unitPrice = t.GetColumn("Уд. цена на 1 кв. м ОКС");
        var sellPriceZu = t.GetColumn("Цена продажи ОКС (ЗУ)");
        var onePrice = t.GetColumn("Цена 1 кв.м ОКС");
        var factUsage = t.GetColumn("Факт. использование ОКС");
        var kadNumber = t.GetColumn("Кад. номер ОКС");
        var area = t.GetColumn("Площадь ОКС");

        foreach (int y in EachRow())
        {
            bool filter = assert.IsEqual(type, "ZU") || assert.IsEqual(type, "Rnt");
            if (!filter) continue;

            filter = assert.IsNotEqual(additionalFilter, "АРХИВ");
            if (!filter) continue;

            filter = assert.IsEmpty(uni);
            if (!filter) continue;

            if (assert.IsNotZeroOrEmpty(sellPriceLot))
            {
                assert.Equal(priceV, "V");

                if (assert.IsEmpty(sellPriceZu) || assert.IsEmpty(sellPriceLot))
                {
                    result.Mark(sellPriceZu, y, $"Ожидалось, что столбцы '{sellPriceZu.name}' и '{sellPriceLot.name}' будут содержать числа, но получены '{t.At<string>(sellPriceZu, y)}' и '{t.At<string>(sellPriceLot, y)}'");
                }
                else
                {
                    float sellPriceZu_value = t.AtFloat(sellPriceZu, y);
                    float sellPriceLot_value = t.AtFloat(sellPriceLot, y);

                    bool mayBeZero = sellPriceZu_value <= LOW_PRICE || sellPriceLot_value <= LOW_PRICE;
                    if (mayBeZero == false)
                    {
                        assert.NotZeroOrEmpty(unitPrice);
                        assert.NotZeroOrEmpty(sellPriceZu);
                        assert.NotZeroOrEmpty(onePrice);
                        assert.NotZeroOrEmpty(factUsage);
                        assert.NotZeroOrEmpty(kadNumber);
                        assert.NotZeroOrEmpty(area);
                    }
                }
            }
            else
            {
                assert.Empty(priceV);

                assert.ZeroOrEmpty(sellPriceZu);
                assert.ZeroOrEmpty(onePrice);

                if (assert.IsNotZeroOrEmpty(unitPrice))
                {
                    assert.NotZeroOrEmpty(kadNumber);
                    assert.NotZeroOrEmpty(area);
                    assert.NotZeroOrEmpty(factUsage);
                }
                else
                {
                    assert.ZeroOrEmpty(kadNumber);
                    assert.ZeroOrEmpty(area);
                    assert.ZeroOrEmpty(factUsage);
                }
            }
        }
    }

    [Test("ЗУ и Rnt (8.)", "ЗУ")]
    private void ZuRnt3()
    {
        var type = t.GetColumn("Вид ОН");
        var additionalFilter = t.GetColumn("Дополн. фильтр");
        var uni = t.GetColumn("Единый лот");

        var unitPrice = t.GetColumn("Уд. цена на 1 кв. м ОКС");
        var sellPriceZu = t.GetColumn("Цена продажи ОКС (ЗУ)");
        var sellPriceLot = t.GetColumn("Цена продажи ОКС (лот)");
        var onePrice = t.GetColumn("Цена 1 кв.м ОКС");
        var factUsage = t.GetColumn("Факт. использование ОКС");
        var area = t.GetColumn("Площадь ОКС");
        var kadNumber = t.GetColumn("Кад. номер ОКС");
        var priceV = t.GetColumn("Стоимость ОКС указана отдельно");

        foreach (int y in EachRow())
        {
            bool filter = assert.IsEqual(type, "ZU") || assert.IsEqual(type, "Rnt");
            if (!filter) continue;

            filter = assert.IsNotEqual(additionalFilter, "АРХИВ");
            if (!filter) continue;

            filter = assert.IsEmpty(uni);
            if (!filter) continue;

            if (assert.IsZeroOrEmpty(unitPrice))
            {
                assert.ZeroOrEmpty(sellPriceZu);
                assert.ZeroOrEmpty(sellPriceLot);
                assert.ZeroOrEmpty(onePrice);
                assert.ZeroOrEmpty(factUsage);
                assert.ZeroOrEmpty(area);
                assert.ZeroOrEmpty(kadNumber);
                assert.ZeroOrEmpty(priceV);
            }
            else
            {
                if (assert.IsNotZeroOrEmpty(sellPriceZu) && assert.IsNotZeroOrEmpty(sellPriceLot))
                {
                    if (assert.IsEmpty(sellPriceZu) || assert.IsEmpty(sellPriceLot))
                    {
                        result.Mark(sellPriceZu, y, $"Ожидалось, что столбцы '{sellPriceZu.name}' и '{sellPriceLot.name}' будут содержать числа, но получены '{t.At<string>(sellPriceZu, y)}' и '{t.At<string>(sellPriceLot, y)}'");
                    }
                    else
                    {
                        float sellPriceZu_value = t.AtFloat(sellPriceZu, y);
                        float sellPriceLot_value = t.AtFloat(sellPriceLot, y);

                        bool mayBeZero = sellPriceZu_value <= LOW_PRICE || sellPriceLot_value <= LOW_PRICE;
                        if (mayBeZero == false)
                        {
                            assert.NotZeroOrEmpty(factUsage);
                            assert.NotZeroOrEmpty(area);
                            assert.NotZeroOrEmpty(kadNumber);
                            assert.NotZeroOrEmpty(onePrice);
                            assert.NotZeroOrEmpty(priceV);
                        }
                    }
                }
                else if (assert.IsZeroOrEmpty(sellPriceZu) && assert.IsZeroOrEmpty(sellPriceLot))
                {
                    assert.NotZeroOrEmpty(factUsage);
                    assert.NotZeroOrEmpty(area);
                    assert.NotZeroOrEmpty(kadNumber);
                    assert.ZeroOrEmpty(onePrice);
                    assert.ZeroOrEmpty(priceV);
                }
                else
                {
                    result.Mark(sellPriceZu, y, $"Исключительный случай: f('{sellPriceZu.name}') != f('{sellPriceLot.name}')");
                }
            }
        }
    }

    [Test("ЗУ и Rnt (9.)", "ЗУ")]
    private void ZuRnt4()
    {
        var type = t.GetColumn("Вид ОН");
        var additionalFilter = t.GetColumn("Дополн. фильтр");
        var uni = t.GetColumn("Единый лот");

        var onePrice = t.GetColumn("Цена 1 кв.м ОКС");
        var unitPrice = t.GetColumn("Уд. цена на 1 кв. м ОКС");
        var sellPriceZu = t.GetColumn("Цена продажи ОКС (ЗУ)");
        var sellPriceLot = t.GetColumn("Цена продажи ОКС (лот)");
        var priceV = t.GetColumn("Стоимость ОКС указана отдельно");
        var kadNumber = t.GetColumn("Кад. номер ОКС");
        var factUsage = t.GetColumn("Факт. использование ОКС");
        var area = t.GetColumn("Площадь ОКС");

        foreach (int y in EachRow())
        {
            bool filter = assert.IsEqual(type, "ZU") || assert.IsEqual(type, "Rnt");
            if (!filter) continue;

            filter = assert.IsNotEqual(additionalFilter, "АРХИВ");
            if (!filter) continue;

            filter = assert.IsEmpty(uni);
            if (!filter) continue;

            if (assert.IsEmpty(onePrice))
            {
                assert.ZeroOrEmpty(sellPriceZu);
                assert.ZeroOrEmpty(sellPriceLot);
                assert.ZeroOrEmpty(priceV);

                if (assert.IsNotEmpty(unitPrice))
                {
                    assert.NotZeroOrEmpty(kadNumber);
                    assert.NotZeroOrEmpty(factUsage);
                    assert.NotZeroOrEmpty(area);
                }
                else
                {
                    assert.ZeroOrEmpty(kadNumber);
                    assert.ZeroOrEmpty(factUsage);
                    assert.ZeroOrEmpty(area);
                }
            }
            else
            {
                if (assert.IsEmpty(sellPriceZu) || assert.IsEmpty(sellPriceLot))
                {
                    result.Mark(sellPriceZu, y, $"Ожидалось, что столбцы '{sellPriceZu.name}' и '{sellPriceLot.name}' будут содержать числа, но получены '{t.At<string>(sellPriceZu, y)}' и '{t.At<string>(sellPriceLot, y)}'");
                }
                else
                {
                    float sellPriceZu_value = assert.IsEmpty(sellPriceZu) ? 0 : t.AtFloat(sellPriceZu, y);
                    float sellPriceLot_value = assert.IsEmpty(sellPriceLot) ? 0 : t.AtFloat(sellPriceLot, y);

                    bool mayBeZero = sellPriceZu_value <= LOW_PRICE || sellPriceLot_value <= LOW_PRICE;
                    if (mayBeZero == false)
                    {
                        assert.NotZeroOrEmpty(sellPriceZu);
                        assert.NotZeroOrEmpty(sellPriceLot);
                        assert.NotZeroOrEmpty(unitPrice);
                        assert.NotZeroOrEmpty(factUsage);
                        assert.NotZeroOrEmpty(area);
                        assert.NotZeroOrEmpty(kadNumber);
                        assert.NotZeroOrEmpty(priceV);
                    }
                }
            }
        }
    }

    [Test("ЕНК", "ЕНК")]
    private void Enk()
    {
        var type = t.GetColumn("Вид ОН");
        var additionalFilter = t.GetColumn("Дополн. фильтр");
        var uni = t.GetColumn("Единый лот");

        var sellPriceZu = t.GetColumn("Цена продажи ЗУ");
        var unitPriceZu = t.GetColumn("Уд. цена на  1 кв.м ЗУ");
        var unitPriceOks = t.GetColumn("Уд. цена на 1 кв. м ОКС");
        var areaOks = t.GetColumn("Площадь ОКС");
        var sellPriceOksLot = t.GetColumn("Цена продажи ОКС (лот)");
        var sellPriceOksZu = t.GetColumn("Цена продажи ОКС (ЗУ)");
        var onePriceOks = t.GetColumn("Цена 1 кв.м ОКС");
        var priceV = t.GetColumn("Стоимость ОКС указана отдельно");
        var kadNumber = t.GetColumn("Кад. номер ОКС");

        foreach (int y in EachRow())
        {
            bool filter = assert.IsEqual(type, "ENC");
            if (!filter) continue;

            filter = assert.IsNotEqual(additionalFilter, "АРХИВ");
            if (!filter) continue;

            filter = assert.IsEmpty(uni);
            if (!filter) continue;

            assert.NotEmpty(sellPriceZu);
            assert.NotEmpty(unitPriceZu);

            if (assert.IsEmpty(unitPriceOks))
            {
                if (! (assert.IsEqual(additionalFilter, "Фиктивный КН") && assert.IsZeroOrEmpty(areaOks)))
                {
                    string addFlt_value = t.At<string>(additionalFilter, y);
                    string areaOks_value = t.At<string>(areaOks, y);
                    result.Mark(unitPriceOks, y, $"Колонка '{unitPriceOks.name}' пуста, а '{additionalFilter.name}' = '{addFlt_value}' и '{areaOks.name}' = '{areaOks_value}'");
                }
            }
            else
            {
                if (assert.IsEqual(additionalFilter, "Фиктивный КН") == false)
                {
                    assert.NotZeroOrEmpty(areaOks, actual => $"Для НЕ 'Фиктивный КН' ожидалось, что '{areaOks.name}' будет НЕ равна 0, но получено '{actual}'");
                }
            }

            assert.Empty(sellPriceOksLot);
            assert.Empty(sellPriceOksZu);
            assert.Empty(onePriceOks);
            assert.Empty(priceV);
            assert.NotEmpty(kadNumber);
        }
    }

    [Test("ОКС", "ОКС")]
    private void Oks()
    {
        var type = t.GetColumn("Вид ОН");
        var additionalFilter = t.GetColumn("Дополн. фильтр");
        var uni = t.GetColumn("Единый лот");

        var sellPriceZu = t.GetColumn("Цена продажи ЗУ");
        var unitPriceZu = t.GetColumn("Уд. цена на  1 кв.м ЗУ");
        var sellPriceLot = t.GetColumn("Цена продажи ЛОТ");
        var sellPriceOksLot = t.GetColumn("Цена продажи ОКС (лот)");
        var sellPriceOksZu = t.GetColumn("Цена продажи ОКС (ЗУ)");
        var unitPriceOks = t.GetColumn("Уд. цена на 1 кв. м ОКС");
        var areaLot = t.GetColumn("Площадь лота");
        var areaOks = t.GetColumn("Площадь ОКС");
        var factUsage = t.GetColumn("Факт. использование ОКС");
        var fractZu = t.GetColumn("Доля ЗУ");
        var areaFractZu = t.GetColumn("Площадь доли ЗУ");
        var priceV = t.GetColumn("Стоимость ОКС указана отдельно");

        foreach (int y in EachRow())
        {
            bool filter = assert.IsEqual(type, "OKS");
            if (!filter) continue;

            filter = assert.IsNotEqual(additionalFilter, "АРХИВ");
            if (!filter) continue;

            filter = assert.IsEmpty(uni);
            if (!filter) continue;


            assert.ZeroOrEmpty(sellPriceZu);
            assert.ZeroOrEmpty(unitPriceZu);


            float sellPriceLot_value = t.AtFloat(sellPriceLot, y);
            float sellPriceOksLot_value = t.AtFloat(sellPriceOksLot, y);
            float sellPriceOksZu_value = t.AtFloat(sellPriceOksZu, y);

            if (! (sellPriceLot_value == sellPriceOksLot_value && sellPriceOksLot_value == sellPriceOksZu_value))
            {
                result.Mark(sellPriceLot, y, "Цены продажи лота, ОКС (лот) и ОКС (ЗУ) не совпадают");
            }

            assert.NotZeroOrEmpty(sellPriceLot);
            assert.NotZeroOrEmpty(sellPriceOksLot);
            assert.NotZeroOrEmpty(sellPriceOksZu);

            assert.NotZeroOrEmpty(unitPriceOks);
            assert.NotZeroOrEmpty(areaLot);
            assert.NotZeroOrEmpty(areaOks);
            assert.NotZeroOrEmpty(factUsage);
            assert.ZeroOrEmpty(fractZu);
            assert.NotZeroOrEmpty(areaFractZu);
            assert.Empty(priceV);
        }
    }
}