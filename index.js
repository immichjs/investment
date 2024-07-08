const xlsx = require('xlsx');

function calculateInvestmentFII(
	initialInvestment = 0,
	monthlyInvestment = 0,
	annualAdjustment = 0,
	quotaPrice = 0,
	lastDividend = 0,
	monthlyPeriod,
) {
	let factor = annualAdjustment / 100;
	let table = [];
	let reinvested = 0;
	let invested = initialInvestment;

	if (!monthlyPeriod) {
		return 'É preciso indicar um periodo em meses.';
	}

	let patrimony = initialInvestment;
	let quotes = 0;
	let dividendMontly = 0;

	for (let i = 1; i <= monthlyPeriod; i++) {
		if (i % 12 === 0) {
			monthlyInvestment *= 1 + factor;
		}

		dividendMontly = quotes * lastDividend;
		invested += monthlyInvestment;
		patrimony += monthlyInvestment + dividendMontly;
		reinvested += dividendMontly;
		quotes = patrimony / quotaPrice;

		table.push({
			invested: +invested,
			patrimony: +patrimony,
			quotes: Math.trunc(quotes),
			monthlyInvestment: +monthlyInvestment,
			monthlyInvestmentWithDividend: +monthlyInvestment + +dividendMontly,
			reinvested: +reinvested,
			dividendMontly: +quotes * +lastDividend,
		});
	}

	return table;
}

const investment = calculateInvestmentFII(11850, 850000, 16, 9.63, 0.11, 120);

const worksheet = xlsx.utils.json_to_sheet(investment);

// Colunas que devem ter formatação de moeda
const colunasMoeda = [
	'invested',
	'patrimony',
	'monthlyInvestment',
	'monthlyInvestmentWithDividend',
	'reinvested',
	'dividendMontly',
];

// Mapeamento dos cabeçalhos para índices de colunas
const range = xlsx.utils.decode_range(worksheet['!ref']);
const headerRow = range.s.r;

// Obtém a posição das colunas que precisam de formatação de moeda
const moedaColIndices = [];
for (let C = range.s.c; C <= range.e.c; ++C) {
	const cellRef = xlsx.utils.encode_cell({ c: C, r: headerRow });
	const cell = worksheet[cellRef];
	if (cell && colunasMoeda.includes(cell.v)) {
		moedaColIndices.push(C);
	}
}

// Aplica a formatação de moeda para as colunas especificadas
for (let R = range.s.r + 1; R <= range.e.r; ++R) {
	moedaColIndices.forEach((C) => {
		const cellRef = xlsx.utils.encode_cell({ c: C, r: R });
		const cell = worksheet[cellRef];
		if (cell && typeof cell.v === 'number') {
			cell.z = 'R$ #,##0.00';
		}
	});
}

const workbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(workbook, worksheet, 'Investimentos');

// Escreve o arquivo no sistema
xlsx.writeFile(workbook, 'investment.xlsx');
