namespace Contract.Api.Models;

public record AgreementRequest(
    string AgreementNumber,
    List<AgreementItemDto> Items
);

public record AgreementItemDto(
    string Name,
    decimal Quantity,
    decimal PriceNoVat,   // цена за единицу БЕЗ НДС
    string Unit = "шт"    // по умолчанию "шт"
);
