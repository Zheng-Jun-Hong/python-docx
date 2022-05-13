from apscheduler.schedulers.blocking import BlockingScheduler
import mongo as m           #報表製作程式
import datetime
import os

def find_all_id(pv):
    plant_id = []
    plant_name = []
    lgroup_id = []
    group_id = []
    plants = pv["plant"]
    equipments = pv["equipment"]
    for plant in plants.find():
        plant_id.append(str(plant["_id"]))
        plant_name.append(plant["name"])
    for name in plant_name:
        lgroups = equipments.find(
            {
                "PV": name,
                "type": "pv_lgroup"
            }
        )
        groups = equipments.find(
            {
                "PV": name,
                "type": "pv_group"
            }
        )
        if lgroups:
            for lgroup in lgroups:
                lgroup_id.append(str(lgroup["_id"]))
        if groups:
            for group in groups:
                group_id.append(str(group["_id"]))
    return plant_id, lgroup_id, group_id

def report_produce_15min(pv, ID, start_time, end_time, logo_path, path=None):
    try:
        plant_id, solar_ID, meter_ID, pr_ID, collection = m.id_identify(pv, ID)
        name, field_position, capacity = m.field_imformation(pv, plant_id)
        time_list = m.set_time_interval(start_time, end_time, "15min")

        irrh_data = m.irrh_cal(pv, solar_ID, time_list, "15min")
        meter_data = m.meter_cal(pv, meter_ID, time_list, "15min")
        pr_data = m.pr_cal(pv, pr_ID, time_list, "15min")
        #有平均值
        # max_val, avg_val = m.avg_max_value(pv, pr_ID, collection, time_list[0], time_list[-1], avg_status=True)
        #無平均值
        max_val = m.avg_max_value(pv, pr_ID, collection, time_list[0], time_list[-1], avg_status=False)
        header_data = m.company_imformation(pv, plant_id)

        name = m.project_name(pv, pr_ID, "15min")
        date = str(start_time) + "~" + str(end_time)
        #有平均值
        # imformation_dict = m.imformation_data(name, date, field_position, str(capacity), str(max_val), str(avg_val))
        #無平均值
        imformation_dict = m.imformation_data(name, date, field_position, str(capacity), str(max_val))
        table_head_datas = ["時間", "日照量", "發電量(kWh)", "PR"]
        data = m.table_data(time_list, irrh_data, meter_data, pr_data)
        m.report(pv, name, header_data, imformation_dict, table_head_datas, data, ID, "15min", logo_path, path=path)
    except Exception as e:
        print(e)

def report_produce_hour(pv, ID, start_time, end_time, logo_path, path=None):
    try:
        plant_id, solar_ID, meter_ID, pr_ID, collection = m.id_identify(pv, ID)
        name, field_position, capacity = m.field_imformation(pv, plant_id)
        time_list = m.set_time_interval(start_time, end_time, "hour")

        irrh_data = m.irrh_cal(pv, solar_ID, time_list, "hour")
        meter_data = m.meter_cal(pv, meter_ID, time_list, "hour")
        pr_data = m.pr_cal(pv, pr_ID, time_list, "hour")
        #有平均值
        # max_val, avg_val = m.avg_max_value(pv, pr_ID, collection, time_list[0], time_list[-1], avg_status=True)
        #無平均值
        max_val = m.avg_max_value(pv, pr_ID, collection, time_list[0], time_list[-1], avg_status=False)
        header_data = m.company_imformation(pv, plant_id)

        name = m.project_name(pv, pr_ID, "hour")
        date = str(start_time) + "~" + str(end_time)
        #有平均值
        # imformation_dict = m.imformation_data(name, date, field_position, str(capacity), str(max_val), str(avg_val))
        #無平均值
        imformation_dict = m.imformation_data(name, date, field_position, str(capacity), str(max_val))
        table_head_datas = ["時間", "日照量", "發電量(kWh)", "PR"]
        data = m.table_data(time_list, irrh_data, meter_data, pr_data)
        m.report(pv, name, header_data, imformation_dict, table_head_datas, data, ID, "hour", logo_path, path=path)
    except Exception as e:
        print(e)

def report_produce_day(pv, ID, start_time, end_time, logo_path, path=None, alarm_data=None, dispatch_data=None):
    try:
        plant_id, solar_ID, meter_ID, pr_ID, collection = m.id_identify(pv, ID)
        name, field_position, capacity = m.field_imformation(pv, plant_id)
        time_list = m.set_time_interval(start_time, end_time, "day")

        irrh_data = m.irrh_cal(pv, solar_ID, time_list, "day")
        meter_data = m.meter_cal(pv, meter_ID, time_list, "day")
        pr_data = m.pr_cal(pv, pr_ID, time_list, "day")
        #有平均值
        # max_val, avg_val = m.avg_max_value(pv, pr_ID, collection, time_list[0], time_list[-1], avg_status=True)
        #無平均值
        max_val = m.avg_max_value(pv, pr_ID, collection, time_list[0], time_list[-1], avg_status=False)
        
        header_data = m.company_imformation(pv, plant_id)

        name = m.project_name(pv, pr_ID, "day")
        date = str(start_time) + "~" + str(end_time)
        #有平均值
        # imformation_dict = m.imformation_data(name, date, field_position, str(capacity), str(max_val), str(avg_val))
        #無平均值
        imformation_dict = m.imformation_data(name, date, field_position, str(capacity), str(max_val))

        table_head_datas = ["時間", "日照量", "發電量(kWh)", "PR"]
        data = m.table_data(time_list, irrh_data, meter_data, pr_data)
        m.report(pv, name, header_data, imformation_dict, table_head_datas, data, ID, "day", logo_path, path=path, alarm_data=alarm_data, dispatch_data=dispatch_data)
    except Exception as e:
        print(e)

def report_produce_month(pv, ID, start_time, end_time, logo_path, path=None):
    try:
        plant_id, solar_ID, meter_ID, pr_ID, collection = m.id_identify(pv, ID)
        name, field_position, capacity = m.field_imformation(pv, plant_id)
        time_list = m.set_time_interval(start_time, end_time, "month")

        irrh_data = m.irrh_cal(pv, solar_ID, time_list, "month")
        meter_data = m.meter_cal(pv, meter_ID, time_list, "month")
        pr_data = m.pr_cal(pv, pr_ID, time_list, "month")
        #有平均值
        # max_val, avg_val = m.avg_max_value(pv, pr_ID, collection, time_list[0], time_list[-1], avg_status=True)
        #無平均值
        max_val = m.avg_max_value(pv, pr_ID, collection, time_list[0], time_list[-1], avg_status=False)

        header_data = m.company_imformation(pv, plant_id)

        name = m.project_name(pv, pr_ID, "month")
        date = str(start_time) + "~" + str(end_time)
        #有平均值
        # imformation_dict = m.imformation_data(name, date, field_position, str(capacity), str(max_val), str(avg_val))
        #無平均值
        imformation_dict = m.imformation_data(name, date, field_position, str(capacity), str(max_val))
        table_head_datas = ["時間", "日照量", "發電量(kWh)", "PR"]
        data = m.table_data(time_list, irrh_data, meter_data, pr_data)
        m.report(pv, name, header_data, imformation_dict, table_head_datas, data, ID, "month", logo_path, path=path)
    except Exception as e:
        print(e)

def calender_to_report():
    path = os.getenv('solar_static', None)
    if path:
        logo_path = path + "/images/logo.jpg"
    else:
        logo_path = "images/logo.jpg"
    pv = m.mongo_connect()
    plant_id, lgroup_id, group_id = find_all_id(pv)
    total_id = plant_id + lgroup_id + group_id
    today = datetime.date.today()
    delta = datetime.timedelta(days=1)
    yesterday = today - delta
    first_day_year = today.replace(month=1, day=1)
    first_day_month = today.replace(day=1)

    request_dict = {
        "time": {
            "start_date": datetime.datetime.strftime(first_day_month, '%Y-%m-%d'),
            "end_date": datetime.datetime.strftime(yesterday, '%Y-%m-%d'),
            "mode": "interval"
        },
        "plant": {"all": False, "ID": [], "col": []},
        "alarm_type": "all",
        "alarm_group": "all",
        "equip_type": "all",
        "page": 1,
        "document_per_page": 999999
    }

    request_dict_dispatch = {"plant": {"ID": "", "col": ""}}

    for id in total_id:
        alarm_data = {"data": []}
        alarm_table = {"data": {"table_data": []}}
        a, solar_ID, meter_ID, pr_ID, collection = m.id_identify(pv, id)
        request_dict["plant"]["ID"] = [id]
        request_dict["plant"]["col"] = [collection]
        alarm_data = m.alarm_get(pv, request_dict)
        alarm_table = m.alarm_table_data(alarm_data["data"])
        alarm_table_data = m.alarm_table_list(alarm_table["data"]["table_data"])

        request_dict_dispatch["plant"]["ID"] = id
        request_dict_dispatch["plant"]["col"] = collection
        dispatch_init = m.get_dispatch_init(pv, request_dict_dispatch)
        dispatch_finish = m.get_dispatch_finish(pv, request_dict_dispatch)

        report_produce_15min(pv, id, yesterday, yesterday, logo_path, path=path)
        report_produce_hour(pv, id, yesterday, yesterday, logo_path, path=path)
        report_produce_day(pv, id, first_day_month, yesterday, logo_path, path=path, alarm_data=alarm_table_data, dispatch_data=[dispatch_finish, dispatch_init])
        report_produce_month(pv, id, first_day_year, yesterday, logo_path, path=path)


if __name__ == "__main__":
    # scheduler = BlockingScheduler(timezone="Asia/Shanghai")
    # scheduler = BlockingScheduler()
    # scheduler.add_job(calender_to_report, 'cron', hour=2, minute=30)
    # scheduler.start()
    calender_to_report()