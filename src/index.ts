import axios, { AxiosResponse } from "axios";
import chalk from "chalk";
import cheerio from "cheerio";
import fs from "fs";
import { Agent } from "https";
import xlsx from "xlsx";

const URL = "https://www.exportargentina.org.ar/companies/";

const httpsAgent = new Agent({ rejectUnauthorized: false });

enum CATEGORIES_ID {
    AGRO = "1",
    INDUSTRIA = "2",
}

let agroPages = 0;
let industryPages = 0;

const promiseSpinner = async (text: string, url: string, params?: any, addLine = true) => {
    let x = 0;
    const P = ["\\", "|", "/", "-"];
    process.stdout.clearLine(0);
    const interval = setInterval(() => {
        process.stdout.write(chalk.grey(`\r${text} ${P[x++]}`));
        x &= 3;
    }, 250);
    let result: AxiosResponse<any>;
    try {
        result = await axios.get(url, { httpsAgent, params });
        process.stdout.write(chalk.grey(`\r${text} `) + chalk.green(`Done!${addLine ? "\n" : ""}`));
    } catch (err) {
        throw new Error(err);
    } finally {
        clearInterval(interval);
    }
    return result;
};

(async () => {
    console.log(chalk.bgCyan("========== EXPORTARGENTINA WEB SCRAPING =========="));
    try {
        const agroCompanies = await promiseSpinner("Obteniendo cantidad de páginas del sector AGRO", URL, { category_id: CATEGORIES_ID.AGRO });
        const agro$ = cheerio.load(agroCompanies.data);
        const agroChildren = agro$(".pagination").children();
        agroPages = Number(agroChildren.filter((i, el) => (i === agroChildren.length - 2)).text());
        console.log(chalk.yellow(`Cantidad de páginas sector AGRO: ${agroPages}`));
    } catch (err) {
        console.log(chalk.red("ERROR Obteniendo cantidad de páginas del sector AGRO"));
    }
    try {
        const industryCompanies = await promiseSpinner("Obteniendo cantidad de páginas del sector INDUSTRIA", URL, { category_id: CATEGORIES_ID.INDUSTRIA });
        const industry$ = cheerio.load(industryCompanies.data);
        const industryChildren = industry$(".pagination").children();
        industryPages = Number(industryChildren.filter((i, el) => (i === industryChildren.length - 2)).text());
        console.log(chalk.yellow(`Cantidad de páginas sector INDUSTRIA: ${industryPages}`));
    } catch (err) {
        console.log(chalk.red("ERROR Obteniendo cantidad de páginas del sector INDUSTRIA"));
    }
    if (agroPages > 0 && industryPages > 0) {
        const companiesId: string[] = [];
        for (let page = 1; page <= agroPages; page++) {
            try {
                const promiseMessage = `Obteniendo ids de empresas sector AGRO página (${Number(page) < 10 ? "0" + page : page} / ${agroPages})`;
                const { data } = await promiseSpinner(promiseMessage, URL, { page, category_id: CATEGORIES_ID.AGRO }, page == agroPages ? true : false);
                const $ = cheerio.load(data);
                $(".panel").map((i, el) => {
                    const companyLink = $(el).parent().attr("href");
                    if (companyLink) companiesId.push(companyLink.split(URL)[1]);
                });
            } catch (err) {
                console.log(chalk.red(`ERROR Obteniendo ids de empresas sector AGRO página: ${page}`));
            }
        }
        const agroCompanies = companiesId.length;
        console.log(chalk.yellow(`Cantidad de empresas sector AGRO: ${agroCompanies}`));
        for (let page = 1; page <= industryPages; page++) {
            try {
                const promiseMessage = `Obteniendo ids de empresas sector INDUSTRIA página (${Number(page) < 10 ? "0" + page : page} / ${industryPages})`;
                const { data } = await promiseSpinner(promiseMessage, URL, { page, category_id: CATEGORIES_ID.INDUSTRIA }, page == industryPages ? true : false);
                const $ = cheerio.load(data);
                $(".panel").map((i, el) => {
                    const companyLink = $(el).parent().attr("href");
                    if (companyLink) companiesId.push(companyLink.split(URL)[1]);
                });
            } catch (err) {
                console.log(chalk.red(`ERROR Obteniendo ids de empresas sector INDUSTRIA página: ${page}`));
            }
        }
        const industryCompanies = companiesId.length - agroCompanies;
        console.log(chalk.yellow(`Cantidad de empresas sector INDUSTRIA: ${industryCompanies}`));
        console.log(chalk.cyan(`Cantidad TOTAL de empresas: ${companiesId.length}`));
        fs.writeFileSync("companies_id.json", JSON.stringify(companiesId), { encoding: "utf8" });
        const companiesIdFile: string[] = JSON.parse(fs.readFileSync("companies_id.json", { encoding: "utf8" }));
        let companies: any[] = [];
        try {
            await Promise.all(companiesIdFile.map((id, index) => {
                return new Promise((resolve, reject) => {
                    setTimeout(async () => {
                        try {
                            process.stdout.clearLine(0);
                            const message = `\rObteniendo datos de la empresa con id ${id} (${index + 1} / ${companiesIdFile.length}) `;
                            process.stdout.write(chalk.grey(message));
                            const { data } = await axios.get(URL + id, { httpsAgent });
                            process.stdout.write(chalk.grey(message) + chalk.green(index + 1 === companiesIdFile.length ? `Done!\n` : `Done!`));
                            const $ = cheerio.load(data);
                            let nombre = $(".section-title").text();
                            let categoria = $(".label-success").text();
                            let descripcion = $("div.margin-20.word-break").text();
                            let mercados = $(".label-primary").map((i, el) => ($(el).text())).get().join(", ");
                            let emailContacto = $("#contact_email").text();
                            let telefono = $("#contact_phone").text().trim();
                            let nombreContacto = "";
                            let cargoContacto = "";
                            let domicilio = "";
                            let provincia = "";
                            let paginaWeb = "";
                            const datosContacto = $("div.media > div.media-body > p.text-muted").map((i, el) => {
                                try {
                                    JSON.parse($(el).text());
                                } catch (err) {
                                    return el;
                                }
                            });
                            for (let i = 0; i < datosContacto.length; i++) {
                                const contacto = datosContacto.filter((index, el) => (index === i)).text().trim().split("\n");
                                if (contacto.length > 1) {
                                    nombreContacto = contacto[0];
                                    cargoContacto = contacto[1].trim();
                                    domicilio = datosContacto.filter((index, el) => (index === i + 1)).text().trim();
                                    provincia = domicilio.split(", ").reverse()[1];
                                    paginaWeb = datosContacto.filter((index, el) => (index === i + 3)).text().trim();
                                    break;
                                }
                            }
                            companies.push({
                                "Nombre": nombre,
                                "Categoría": categoria,
                                "Descripción": descripcion,
                                "Mercados": mercados,
                                "Nombre Contacto": nombreContacto,
                                "Cargo Contacto": cargoContacto,
                                "Email Contacto": emailContacto,
                                "Domicilio": domicilio,
                                "Provincia": provincia,
                                "Página Web": paginaWeb,
                                "Teléfono": telefono,
                            });
                            resolve(null);
                        } catch (err) {
                            process.stdout.clearLine(0);
                            process.stdout.write(chalk.grey(`\rObteniendo datos de la empresa con id ${id} (${index + 1} / ${companiesIdFile.length}) `));
                            reject(`\nERROR Obteniendo datos de empresa con id ${id}`);
                        }
                    }, 500 * index);
                });
            }));
        } catch (err) {
            console.log(chalk.red(err));
        }
        try {
            const workbook = xlsx.utils.book_new();
            console.log(chalk.grey("Guardando datos en excel"));
            const worksheet = xlsx.utils.json_to_sheet(companies, { header: [
                "Nombre",
                "Categoría",
                "Descripción",
                "Mercados",
                "Nombre Contacto",
                "Cargo Contacto",
                "Email Contacto",
                "Domicilio",
                "Provincia",
                "Página Web",
                "Teléfono",
            ] });
            xlsx.utils.book_append_sheet(workbook, worksheet, "Hoja 1");
            xlsx.writeFile(workbook, "empresas.xlsx");
            console.log(chalk.green("Datos guardados con éxito!"));
        } catch (err) {
            console.log(chalk.red("ERROR Guardando datos en excel"));
        }
    }
})();
