namespace Contract.Api.Models;

public record AgreementRequest(
    string AgreementNumber,
    string BuyerName,
    string BuyerInn,
    string BuyerAddress,
    string BuyerPhone,
    string BuyerAccount,
    string BuyerBank,
    string BuyerMfo,
    string BuyerDirector,
    List<AgreementItemDto> Items
);

public record AgreementItemDto(
    string Name,
    decimal Quantity,
    decimal PriceNoVat,
    string Unit = "шт"
);

