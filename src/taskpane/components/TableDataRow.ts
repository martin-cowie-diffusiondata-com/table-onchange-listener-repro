import { faker } from "@faker-js/faker";

export class TableDataRow {

    private constructor(
        readonly prefix: string,
        readonly firstName: string,
        readonly familyName: string,
        readonly dob: string,
        readonly jobTitle: string,
        readonly accountNumber: string,
        readonly amount: number,
    ) {}

    public static build() {
        return new TableDataRow(
            faker.person.prefix("male"),
            faker.person.firstName(),
            faker.person.lastName(),
            faker.date.birthdate().toISOString(),
            faker.person.jobTitle(),
            faker.finance.accountNumber(),
            faker.number.int(100_000)
        );
    }
    static columns = ["prefix", "firstName", "familyName", "dob", "jobTitle", "accountNumber", "amount"];
}

