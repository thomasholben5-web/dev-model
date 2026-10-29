"""sample_deal.py — apply a deliberately-different made-up deal to stress-test the model."""
import sys, openpyxl
def build(wb):
    ip=wb['Inputs']; um=wb['UnitMatrix']; tl=wb['Timeline']
    # --- unit mix / rents (150 units, different SF & rents) ---
    um['D4']=40;  um['E4']=700;  um['G4']=1500   # Studio
    um['D5']=60;  um['E5']=1000; um['G5']=1900   # 1BR-A
    um['D7']=50;  um['E7']=1250; um['G7']=2400   # 2BR-A (was 0)
    # --- scalar assumptions ---
    ip['O9']=4000000        # LandCost
    ip['C20']=0.060         # EntryCapRate
    ip['G23']=0.60          # TargetLTC
    ip['G29']=0.075         # ConRate
    ip['G33']=0.085         # ConMinDebtYield
    ip['G34']=1.25          # ConMinDSCR
    ip['C19']=84            # HoldPeriodMonths (7-yr -> sell)
    ip['C30']=0.035         # GrowthRent
    ip['C32']=0.025         # GrowthExpense
    ip['K22']=0.035         # DevFeePctDirect (turn ON)
    ip['C45']=15            # DeliveryLagMonths
    ip['C41']=20            # UnitsLeasedPerMo
    ip['C40']=25            # UnitsDeliveredPerMo
    ip['K38']='Reassessed'  # reassess-on-sale path ON
    # --- timeline durations (direct inputs now) ---
    tl['D4']=3; tl['D5']=6; tl['D6']=2; tl['D7']=24
    return {'units':150}
def main():
    wb=openpyxl.load_workbook(sys.argv[1]); print('applied:',build(wb)); wb.save(sys.argv[2]); print('saved',sys.argv[2])
if __name__=='__main__': main()
