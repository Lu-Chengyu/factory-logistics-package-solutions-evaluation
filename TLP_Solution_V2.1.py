import os, numpy as np, math, decimal as dc, pulp as pp, itertools as its, xlrd, xlwt, xlutils.copy as copy, time
err = 1e-06
tolerance_1 = 2

def OneD_To_TwoD(Llist=[], Wlist=[], LWlist=[]):
    if len(Llist) != len(Wlist):
        print('OneD_To_TwoD Err')
    else:
        print('OneD_To_TwoD transform processing...')
        time.sleep(0.5)
        for i in range(len(Llist)):
            if Llist[i] != '':
                LWlist.append([Llist[i], Wlist[i]])

        print('OneD_To_TwoD transform succeeded')
    return LWlist


def deci(list_0=[]):
    for i, list_item in enumerate(list_0):
        if isinstance(list_item, float):
            list_item = dc.Decimal(str(list_item))
        else:
            if len(list_item) > 1:
                for j in range(len(list_item)):
                    if isinstance(list_item[j], float):
                        list_item[j] = dc.Decimal(str(list_item[j]))

    return list_0


def type_match_unique(type_list=[], match=[]):
    type_temp = []
    match_temp = []
    match = list(match)
    for i in range(len(type_list)):
        if type_list[i] != '' and type_list[i] not in type_temp:
            type_temp.append(type_list[i])
            if type(match[i]).__name__ == 'float':
                match_temp.append(round(match[i], 5))
            else:
                match_temp.append(match[i])

    return (
     type_temp, match_temp)


def type_match_group(type_list=[], match=[], unique=1):
    type_temp = []
    match_temp = []
    match = list(match)
    for i in range(len(type_list)):
        if type_list[i] != '' and type_list[i] not in type_temp:
            type_temp.append(type_list[i])

    for k in range(len(type_temp)):
        match_temp.append([])

    for k in range(len(type_temp)):
        for i in range(len(type_list)):
            if type_temp[k] == type_list[i]:
                if unique == 1:
                    if type(match[i]).__name__ == 'float':
                        if round(match[i], 5) not in match_temp[k]:
                            match_temp[k].append(round(match[i], 5))
                        else:
                            if match[i] not in match_temp[k]:
                                match_temp[k].append(match[i])
                    else:
                        if type(match[i]).__name__ == 'float':
                            match_temp[k].append(round(match[i], 5))
                        else:
                            match_temp[k].append(match[i])

    return (
     type_temp, match_temp)


def type_add(type=[], add=[]):
    type_temp = []
    add_temp = []
    add = list(add)
    for i in range(len(type)):
        if type[i] != '' and type[i] not in type_temp:
            type_temp.append(type[i])

    for k in range(len(type_temp)):
        if isinstance(add[i], list):
            add_temp.append([])
            for j in range(len(add[i])):
                add_temp[k].append(0)

        else:
            add_temp.append(0)

    for k in range(len(type_temp)):
        for i in range(len(type)):
            if type_temp[k] == type[i]:
                if isinstance(add[i], list):
                    for j in range(len(add[i])):
                        add_temp[k][j] += add[i][j]

                else:
                    add_temp[k] += round(add[i], 5)

    return (
     type_temp, add_temp)


def type_count(list_func=[], original_type_list=[], match_type_list=[]):
    list_type = []
    list_type_uni = []
    list_count = []
    for i in range(len(list_func)):
        for j in range(len(match_type_list)):
            if list_func[i][1] == match_type_list[j]:
                list_type.append(original_type_list[j])
                if original_type_list[j] not in list_type_uni:
                    list_type_uni.append(original_type_list[j])

    for i in range(len(list_type_uni)):
        list_count.append(list_type.count(list_type_uni[i]))

    return (list_type_uni, list_count)


def Combo_2D_to_1D(Combo_func):
    combo_temp = []
    position_temp = []
    for i in range(len(Combo_func)):
        position_temp.append(0)
        for j in range(len(Combo_func[i])):
            combo_temp.append(Combo_func[i][j])
            position_temp[i] += 1

    return (
     combo_temp, position_temp)


def List_2D_to_1D(list_func):
    list_temp = []
    for i in range(len(list_func)):
        for j in range(len(list_func[i])):
            list_temp.append(list_func[i][j])

    return list_temp


def List_1D_to_2D(list_func_1D, list_match_2D):
    list_temp = []
    for i in range(len(list_match_2D)):
        list_temp.append([])
        for j in range(len(list_match_2D[i])):
            list_temp[i].append(0)

    num = 0
    for i in range(len(list_match_2D)):
        for j in range(len(list_match_2D[i])):
            list_temp[i][j] = list_func_1D[num]
            num += 1

    return list_temp


def Print_CMD(Nname, Ndescription, Narray):
    print('')
    print(Nname, ':', Ndescription)
    print('box arrangement after ', Nname, ':')
    time.sleep(0.5)
    for i in range(len(Box_H)):
        print('')
        print('height', i + 1, ':\t', Box_H[i])
        if H_Complete[i] == 0:
            print('completed in former step')
        else:
            print('>>>>>> box arranged in ', Nname, ':')
            layer_none = 0
            for j in range(len(Narray)):
                if round(Narray[j][i]) > 0:
                    print('>>>>>>>>>>>> ', round(Narray[j][i]), ' \tlayer of pallet', j + 1, Pallet_LW_avail[j])
                    layer_none = 1

            if layer_none == 0:
                print('>>>>>>>>>>>> none')
            remain_box_none = 0
            print('>>>>>> box remained:')
            for j in range(len(LW_RB_groupby_H[i])):
                if Num_RB_groupby_H[i][j] > 0:
                    print('>>>>>>>>>>>> box dimension:', LW_RB_groupby_H[i][j], '\tnum:', Num_RB_groupby_H[i][j])
                    remain_box_none = 1

        if remain_box_none == 0:
            print('>>>>>>>>>>>> none')
            H_Complete[i] = 0
        if Nname == 'N2':
            print('>>>>>> box to fill in:')
            count_fill_in = 0
            for j in range(len(LW_RB_groupby_H[i])):
                if Num_RB_groupby_H[i][j] < 0:
                    print('>>>>>>>>>>>> box dimension:', LW_RB_groupby_H[i][j], '\tnum:', -Num_RB_groupby_H[i][j])
                    count_fill_in += 1

            if count_fill_in == 0:
                print('>>>>>>>>>>>> none')


def del_rdd(list=[]):
    list_temp = []
    for i, item in enumerate(list):
        if item not in list_temp:
            list_temp.append(item)

    return list_temp


def del_blk(list=[]):
    list_temp = []
    for i, item in enumerate(list):
        if item != '':
            list_temp.append(item)

    return list_temp


def del_blank_queue(*list1, length):
    list_temp = list(list1).copy()
    for k in range(len(list_temp)):
        list_temp[k] = []

    for j, item in enumerate(list1[0]):
        if item != '':
            for i in range(length):
                list_temp[i].append(list1[i][j])

        else:
            for i in range(length):
                list_temp[i].append(-1)

    for i in range(length):
        list_temp[i] = list(filter((lambda x: x != -1), list_temp[i]))

    for i in range(length):
        for j in range(len(list_temp[i])):
            if type(list_temp[i][j]).__name__ == 'float':
                list_temp[i][j] = round(list_temp[i][j], 5)

    return list_temp


def Combo_Filter(Combo_temp, Num_R_box1, buffer=0):
    if Combo_temp == []:
        return Combo_temp
    else:
        Combo_temp_copy = Combo_temp.copy()
        box_no = len(Combo_temp[0])
        for i1 in range(len(Combo_temp_copy)):
            Combo_temp_copy[i1] = list(Combo_temp_copy[i1])
            for i2 in range(len(Combo_temp_copy[i1])):
                if Combo_temp_copy[i1][i2] > Num_R_box1[i2] + buffer:
                    Combo_temp_copy[i1] = -1
                    break

        positive = 0
        for i, c in enumerate(Combo_temp_copy):
            if c != -1:
                positive = 1

        if positive == 1:
            Combo_temp_copy = list(filter((lambda x: x != -1), Combo_temp_copy))
        else:
            Combo_temp_copy = []
        success_once = []
        for i3 in range(box_no):
            success_once.append(0)

        for i4 in range(len(Combo_temp_copy)):
            for i5 in range(len(Combo_temp_copy[i4])):
                if Combo_temp_copy[i4][i5] >= min(1, Num_R_box1[i5]):
                    success_once[i5] = 1

        box_length = []
        for i6 in range(len(Combo_temp)):
            temp = 0
            for i7 in range(len(Combo_temp[i6])):
                temp += Combo_temp[i6][i7]

            box_length.append(temp)

        combo_add = []
        add_l = 0
        for i8 in range(len(success_once)):
            box_l_min = 999
            if success_once[i8] == 0:
                add_l += 1
                for i9, cb in enumerate(Combo_temp):
                    if (cb[i8] >= 1) & (box_length[i9] < box_l_min):
                        if len(combo_add) >= add_l:
                            combo_add.pop()
                            success_once[i8] = 1
                        combo_add.append(cb)

            if i8 + 1 < len(success_once):
                for i10 in range(i8 + 1, len(success_once)):
                    if combo_add != [] and combo_add[-1][i10] >= 1:
                        success_once[i10] = 1

        Combo_temp = Combo_temp_copy + combo_add
        return Combo_temp


def Del_Pallet(array_func, pallet_avail):
    pallet_temp = []
    for i in range(len(array_func)):
        count = 0
        for j in range(len(array_func[i])):
            if array_func[i][j] != 0:
                count += 1

        if count != 0:
            pallet_temp.append(pallet_avail[i])

    return pallet_temp


def type_match_sub(sub_a, group_a, group_b):
    sub_b = []
    for i in range(len(group_a)):
        if group_a[i] in sub_a:
            sub_b.append(group_b[i])

    return sub_b


def arrange_box(some_box=[], border=[], judge_point=[], The_Realm_Area=[], solution=[], deep=0, area=0, some_box_skiped=[], time_start=0):
    solution_max = []
    solution_length = 0
    solution_temp = solution.copy()
    some_box_copy = some_box.copy()
    some_box_temp = some_box.copy()
    max_border = border[-1]
    success_once = 0
    timeout_y = 0
    if time_start == 0:
        time_start = time.time()
    else:
        time_now = time.time()
        if time_now - time_start > timeout_set:
            done = -1
            return (solution, some_box_copy, done, timeout_y)
        if some_box_copy != []:
            a_box = some_box_copy.pop(0)
            done = 0
            area_min = a_box[0] * a_box[1]
            for box in enumerate(some_box_copy):
                if box[1][0] * box[1][1] < area_min:
                    area_min = box[1][0] * box[1][1]

        else:
            if some_box_skiped == []:
                done = 1
            else:
                done = 0
            return (
             solution, some_box_copy, done, timeout_y)
    if The_Realm_Area - area < area_min:
        done = 1
        some_box_copy.insert(0, a_box)
        return (
         solution, some_box_copy, done, timeout_y)
    else:
        for i, corner in enumerate(border):
            if corner == max_border:
                continue
            else:
                solution_temp = solution.copy()
                border_temp = border.copy()
                judge_temp = judge_point.copy()
                point_LT = [
                 corner[0], corner[1] + a_box[1]]
                point_RB = [corner[0] + a_box[0], corner[1]]
                point_LB = corner
                point_RT = [corner[0] + a_box[0], corner[1] + a_box[1]]
                one_step = [point_LB, point_RT]
                area_add = (point_RT[1] - point_LB[1]) * (point_RT[0] - point_LB[0])
                virtue_list = []
                virtue_point_x = 0
                for i, bt in enumerate(border_temp):
                    if bt[0] < point_LT[0]:
                        virtue_list.append(bt[0])

                for j, jt in enumerate(judge_temp):
                    if (jt[1] > point_LT[1]) & (jt[0] in virtue_list) and jt[0] > virtue_point_x:
                        virtue_point_x = jt[0]

                virtue_point = [
                 virtue_point_x, point_LT[1]]
                cross_judge = 0
                for i, step in enumerate(solution):
                    if (point_RT[0] > step[0][0]) & (point_RT[1] > step[0][1]) & (point_LB[0] < step[1][0]) & (point_LB[1] < step[1][1]):
                        cross_judge = 1

                if (point_LT[1] <= max_border[1]) & (point_RB[0] <= max_border[0]) & (cross_judge == 0):
                    for i, bt in enumerate(border_temp):
                        if (bt[0] == point_LT[0]) & (bt[1] > point_LB[1]) & (bt[1] < point_LT[1]):
                            border_temp.remove(bt)

                    if (point_LT[1] < max_border[1]) & (virtue_point not in border_temp):
                        border_temp.append(virtue_point)
                    if point_LB in border_temp:
                        border_temp.remove(point_LB)
                    if point_LT in judge_temp:
                        judge_temp.remove(point_LT)
                    else:
                        if point_LT[1] == max_border[1]:
                            pass
                        else:
                            if point_LT not in border_temp:
                                border_temp.append(point_LT)
                    if point_RB in judge_temp:
                        judge_temp.remove(point_RB)
                    else:
                        if point_RB[0] == max_border[0]:
                            pass
                        else:
                            if point_RB not in border_temp:
                                border_temp.append(point_RB)
                            solution_temp.append(one_step)
                            if success_once == 0:
                                area += area_add
                            if point_RT in border_temp:
                                border_temp.remove(point_RT)
                            else:
                                judge_temp.append(point_RT)
                if max_border not in border_temp:
                    border_temp.append(max_border)
                border_temp = sorted(border_temp)
                box_memo = some_box_copy
                if (point_LT[1] <= max_border[1]) & (point_RB[0] <= max_border[0]) & (cross_judge == 0):
                    deep += 1
                    success_once = 1
                    solution_memo, box_memo, done, timeout_y = arrange_box(box_memo, border_temp, judge_temp, The_Realm_Area, solution_temp, deep, area, some_box_skiped, time_start)
                    if done == 1:
                        return (
                         solution_memo, box_memo, done, timeout_y)
                    if done == -1:
                        return (
                         solution_memo, box_memo, done, timeout_y)
                    deep -= 1
                else:
                    solution_memo = solution_temp
            if len(solution_memo) > solution_length:
                solution_max = solution_memo
                solution_length = len(solution_memo)
                some_box_temp = box_memo

        if success_once == 0:
            some_box_list1 = []
            some_box_list2 = []
            some_box_list1.append(a_box)
            for i, box in enumerate(some_box_temp):
                pop_box = some_box_temp[i]
                if (a_box[0] == pop_box[0]) & (a_box[1] == pop_box[1]) | (a_box[1] == pop_box[0]) & (a_box[0] == pop_box[1]):
                    some_box_list1.append(pop_box)
                else:
                    some_box_list2.append(pop_box)

            if some_box_list2 == []:
                some_box_temp = some_box_list1
            else:
                deep += 1
                solution_memo, box_memo, done, timeout_y = arrange_box(some_box_list2, border_temp, judge_temp, The_Realm_Area, solution_temp, deep, area, some_box_list1, time_start)
                box_memo += some_box_list1
                if done == 1:
                    return (
                     solution_memo, box_memo, done, timeout_y)
                if done == -1:
                    return (
                     solution_memo, box_memo, done, timeout_y)
                deep -= 1
            if len(solution_memo) > solution_length:
                solution_max = solution_memo
                solution_length = len(solution_memo)
                some_box_temp = box_memo
        return (
         solution_max, some_box_temp, done, timeout_y)


def swerve_by_group(some_box=[]):
    some_box_type = del_rdd(some_box)
    count_type = len(some_box_type)
    queue1 = []
    for num in range(2 ** count_type):
        new_box_list = []
        for i in range(count_type):
            for j in range(len(some_box)):
                if some_box[j] == some_box_type[i]:
                    if 1 << i & num:
                        new_box_list.append(list(reversed(some_box[j])))
                    else:
                        new_box_list.append(list(some_box[j]))

        queue1.append(new_box_list)

    return queue1


def linear_combo(some_box, the_realm):
    some_box_type = del_rdd(some_box)
    count_type = len(some_box_type)
    length_type = List_2D_to_1D(some_box_type)
    length_type = del_rdd(length_type)
    print('pause')


def Box_Master(some_box=[], The_Realm=[], new_box_list=[], The_Judge=[]):
    some_box = deci(some_box)
    The_Realm = deci(The_Realm)
    The_Realm_Area = (The_Realm[1][0] - The_Realm[0][0]) * (The_Realm[1][1] - The_Realm[0][1])
    count = len(some_box)
    length = 0
    queue = []
    time_start = time.time()
    timeout_Y = 0
    queue += swerve_by_group(some_box)
    for i, q in enumerate(queue):
        arrangement, box_list_copy, done, timeout_Y = arrange_box(q, The_Realm, The_Judge, The_Realm_Area)
        if done == 1:
            best_fit = arrangement
            length = len(best_fit)
            box_remain = box_list_copy
            break
        elif arrangement != None:
            pass
        if len(arrangement) > length:
            best_fit = arrangement
            length = len(best_fit)
            box_remain = box_list_copy
        time_now = time.time()
        if time_now - time_start > timeout_set:
            timeout_Y = 1
            break
        else:
            timeout_Y = 0

    if (done == 0) & (timeout_Y == 0):
        for num in range(2 ** count):
            new_box_list = []
            for i in range(count):
                if 1 << i & num:
                    new_box_list.append(list(reversed(some_box[i])))
                else:
                    new_box_list.append(list(some_box[i]))

            if new_box_list in queue:
                continue
            arrangement, box_list_copy, done, timeout_Y = arrange_box(new_box_list, The_Realm, The_Judge, The_Realm_Area)
            if done == 1:
                best_fit = arrangement
                length = len(best_fit)
                box_remain = box_list_copy
                break
            elif arrangement != None:
                pass
            if len(arrangement) > length:
                best_fit = arrangement
                length = len(best_fit)
                box_remain = box_list_copy
            time_now = time.time()
            if time_now - time_start > timeout_set:
                break

    if best_fit != []:
        for i in range(len(best_fit)):
            for j in range(len(best_fit[i])):
                best_fit[i][j][0] = float(best_fit[i][j][0])
                best_fit[i][j][1] = float(best_fit[i][j][1])

    if box_remain != []:
        for i in range(len(box_remain)):
            box_remain[i][0] = float(box_remain[i][0])
            box_remain[i][1] = float(box_remain[i][1])

    for i in range(len(some_box)):
        some_box[i][0] = float(some_box[i][0])
        some_box[i][1] = float(some_box[i][1])

    for i in range(len(The_Realm)):
        The_Realm[i][0] = float(The_Realm[i][0])
        The_Realm[i][1] = float(The_Realm[i][1])

    return (
     best_fit, box_remain)


def P_arrange(Pallet=[], Box=[], mode=0):
    if Box == []:
        return []
    else:
        Max_Permit = []
        Area = []
        for i in range(len(Box)):
            Max_Permit.append(math.ceil(Pallet[0] * Pallet[1] / (Box[i][0] * Box[i][1])))
            Area.append(Box[i][0] * Box[i][1])

        def Combine(Max_Permit=[], n=0, Sol=[], x=np.zeros((len(Max_Permit)), dtype=int)):
            for i in range(Max_Permit[n] + 1):
                x[n] = i
                if n == len(Max_Permit) - 1:
                    Sol.append(tuple(x.copy()))
                else:
                    n += 1
                    Sol = Combine(Max_Permit, n, Sol, x)
                    n -= 1

            return Sol

        Combo1 = Combine(Max_Permit)
        for i in range(len(Combo1)):
            for j in range(len(Box)):
                if Combo1[i][j] > Max_Permit[j] - 1 + mode:
                    Combo1[i] = 0
                    break

            if Combo1[i] == 0:
                continue
            else:
                sum = 0
                for j1 in range(len(Area)):
                    sum += Area[j1] * Combo1[i][j1]

            if abs(sum - Pallet[0] * Pallet[1]) > err:
                Combo1[i] = 0
                continue

        Combo1 = list(filter(None, Combo1))
        return Combo1


def Box_List(Combo=[], Box=[], Sol=[]):
    for con in range(len(Combo)):
        Boxlist = []
        for i in range(len(Box)):
            if Combo[con][i] <= 0:
                continue
            else:
                con1 = 0
                while con1 < Combo[con][i]:
                    Boxlist.append(Box[i])
                    con1 += 1

        Sol.append(Boxlist)

    return Sol


def remove_0(Box_Type=[], Num_Boxes=[], Combo=[]):
    num_box_copy = list.copy(Num_Boxes)
    for j in range(len(Num_Boxes)):
        judge = 0
        for i in range(len(Combo)):
            if Combo[i][j] == 0:
                continue
            else:
                judge = 1
                break

        if judge == 0:
            num_box_copy[j] = 0

    return num_box_copy


def ILP_Box(Combo=[], Num_Boxes_ori=[], Num_Boxes=[]):
    V_num = len(Combo)
    Err = -1
    status = 0
    while status != 1 or Err > 30:
        Err += 1
        prob1 = pp.LpProblem('Box-ILP1', pp.LpMinimize)
        variables_x = []
        variables_u = []
        variables_v = []
        count = V_num
        count2 = len(Num_Boxes_ori)
        if count < 11:
            variables_x += [pp.LpVariable(('X' + str(1) + '%d' % i), lowBound=0, cat=(pp.LpInteger)) for i in range(count)]
        else:
            while count > 0:
                n = int((V_num - count) / 10)
                str_num = 1 * 10 ** n
                for k in range(n):
                    str_num += 9 * 10 ** k

                if count >= 10:
                    variables_x += [pp.LpVariable(('X' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(10)]
                else:
                    variables_x += [pp.LpVariable(('X' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(count)]
                count -= 10

        if count2 < 11:
            variables_u += [pp.LpVariable(('u' + str(1) + '%d' % i), lowBound=0, cat=(pp.LpInteger)) for i in range(count2)]
            variables_v += [pp.LpVariable(('v' + str(1) + '%d' % i), lowBound=0, cat=(pp.LpInteger)) for i in range(count2)]
        else:
            while count2 > 0:
                n = int((len(Num_Boxes_ori) - count2) / 10)
                str_num = 1 * 10 ** n
                for k in range(n):
                    str_num += 9 * 10 ** k

                if count2 >= 10:
                    variables_u += [pp.LpVariable(('u' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(10)]
                    variables_v += [pp.LpVariable(('v' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(10)]
                else:
                    variables_u += [pp.LpVariable(('u' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(count2)]
                    variables_v += [pp.LpVariable(('v' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(count2)]
                count2 -= 10

        variables = variables_x + variables_u + variables_v
        obj = sum(variables[i] for i in range(V_num, 2 * len(Num_Boxes_ori) + V_num))
        prob1 += obj
        for j in range(len(Num_Boxes)):
            if Num_Boxes[j] == 0:
                prob1 += sum(Combo[i][j] * variables[i] for i in range(V_num)) == 0
            else:
                prob1 += sum(Combo[i][j] * variables[i] for i in range(V_num)) - Num_Boxes[j] >= -Err
                prob1 += sum(Combo[i][j] * variables[i] for i in range(V_num)) - Num_Boxes[j] <= Err
                prob1 += sum(Combo[i][j] * variables[i] for i in range(V_num)) - Num_Boxes[j] - variables[V_num + j] + variables[V_num + len(Num_Boxes_ori) + j] == 0

        status = prob1.solve()

    Sol = []
    for i in prob1.variables():
        Sol.append(i.varValue)

    Sol = Sol[0:V_num]
    R_Boxes1 = []
    for i in range(len(Num_Boxes)):
        R_Boxes1.append(int(-sum(Sol[j] * Combo[j][i] for j in range(V_num)) + Num_Boxes_ori[i]))

    return (Sol, R_Boxes1, prob1)


def ILP_Box_nonleft(combo_func, num_rb, num_rb_errproof):
    prob3 = pp.LpProblem('MyPro_N3', pp.LpMinimize)
    total_count = len(combo_func)
    variables = []
    for j in range(len(combo_func)):
        count = len(combo_func[j])
        if count < 11:
            variables += [pp.LpVariable(('P' + str(j + 1) + '%d' % i), lowBound=0, cat=(pp.LpInteger)) for i in range(count)]
        else:
            while count > 0:
                n = int((len(combo_func[j]) - count) / 10)
                str_num = (j + 1) * 10 ** n
                for k in range(n):
                    str_num += 9 * 10 ** k

                if count >= 10:
                    variables += [pp.LpVariable(('P' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(10)]
                else:
                    variables += [pp.LpVariable(('P' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(count)]
                count -= 10

    obj = sum(variables[i] for i in range(total_count))
    prob3 += obj
    for j in range(len(num_rb_errproof)):
        if num_rb_errproof[j] == 0:
            prob3 += sum(combo_func[i][j] * variables[i] for i in range(total_count)) == 0
        else:
            prob3 += sum(combo_func[i][j] * variables[i] for i in range(total_count)) >= num_rb_errproof[j]

    status = prob3.solve()
    var_combo = []
    for i in prob3.variables():
        var_combo.append(i.varValue)

    rb_n3 = []
    for b in range(len(num_rb)):
        rb_n3.append(int(-sum(var_combo[j] * combo_func[j][b] for j in range(total_count)) + num_rb[b]))

    return (var_combo, rb_n3, prob3)


def N3_Generator(var_c, position_c, combo_n3_1d, lw_byH, num_byH_byLW, lw_type, box_h):
    n3_temp = []
    for i in range(len(position_c)):
        n3_temp.append([])
        for j in range(len(box_h)):
            n3_temp[i].append(0)

    h_byLW = []
    for j in range(len(lw_type)):
        h_byLW.append([])
        for i in range(len(box_h)):
            for k in range(len(lw_byH[i])):
                if lw_byH[i][k] == lw_type[j]:
                    CD = num_byH_byLW[i][k]
                    while CD > 0:
                        h_byLW[j].append(box_h[i])
                        CD -= 1

    combo_h = []
    for i in range(len(combo_n3_1d)):
        CD = var_c[i]
        combo_h.append(0)
        while CD > 0:
            for j in range(len(lw_type)):
                for k in range(combo_n3_1d[i][j]):
                    if h_byLW[j] != []:
                        h_temp = h_byLW[j].pop(0)
                        if h_temp > combo_h[i]:
                            combo_h[i] = h_temp

            CD -= 1

    RB_AN3_temp = []
    for i in range(len(h_byLW)):
        if h_byLW[i] != []:
            RB_AN3_temp.append([lw_type[i], h_byLW[i]])

    for i in range(len(Num_RB_groupby_H)):
        for j in range(len(Num_RB_groupby_H[i])):
            Num_RB_groupby_H[i][j] = 0

    combo_h_byP = []
    for i in range(len(position_c)):
        combo_h_byP.append([])
        CD = position_c[i]
        while CD > 0:
            combo_h_byP[i].append(combo_h.pop(0))
            CD -= 1

    for i in range(len(combo_h_byP)):
        for j in range(len(combo_h_byP[i])):
            for h in range(len(box_h)):
                if combo_h_byP[i][j] == box_h[h]:
                    n3_temp[i][h] += 1

    return (
     n3_temp, RB_AN3_temp)


def L_arrange(LimitHeight=1.02, H=[]):
    H = list(reversed(sorted(H)))
    Unsuit_H = []
    for i in range(len(H)):
        if H[i] > LimitHeight or H[i] == 0:
            Unsuit_H.append(H[i])
            H[i] = -1

    H = list(filter((lambda x: x != -1), H))
    N = []
    Layers = []
    for i in range(len(H)):
        N.append(int(LimitHeight / H[i]))

    for i in range(len(N)):
        for j in range(N[i]):
            Layers.append(H[i])

    if N == []:
        return []
    else:
        Combo1 = list(its.combinations(Layers, N[-1]))
        Combo1 = sorted((set(Combo1)), key=(Combo1.index))
        Combo = []
        for i in range(len(Combo1)):
            sum_temp = 0
            Combo_temp = []
            for j in range(N[-1]):
                sum_temp = sum_temp + Combo1[i][j]
                if sum_temp <= LimitHeight:
                    Combo_temp.append(Combo1[i][j])
                elif sum_temp > LimitHeight:
                    Combo_temp.append(0)

            Combo.append(tuple(Combo_temp))

        Combo = sorted((set(Combo)), key=(Combo.index))
        return Combo


def LP_Truck(vol, truck_vol, truck_price, truck_LW, truck_list, buffer_func, price_list=[]):
    for i in range(len(truck_vol)):
        truck_vol[i] = round(truck_vol[i] * (1 - buffer_func), 2)

    prob1 = pp.LpProblem('Truck-ILP1', pp.LpMinimize)
    variables = []
    count = len(truck_vol)
    if count < 11:
        variables += [pp.LpVariable(('T' + str(1) + '%d' % i), lowBound=0, cat=(pp.LpInteger)) for i in range(count)]
    else:
        while count > 0:
            n = int((len(truck_vol) - count) / 10)
            str_num = 1 * 10 ** n
            for k in range(n):
                str_num += 9 * 10 ** k

            if count >= 10:
                variables += [pp.LpVariable(('T' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(10)]
            else:
                variables += [pp.LpVariable(('T' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(count)]
            count -= 10

    obj = sum(variables[i] * truck_price[i] for i in range(len(truck_vol)))
    prob1 += obj
    prob1 += sum(variables[i] * truck_vol[i] for i in range(len(truck_vol))) - vol >= 0
    status = prob1.solve()
    vol_total = 0
    i = -1
    for x in prob1.variables():
        i += 1
        if x.varValue != 0:
            count = x.varValue
            while count > 0:
                truck_list.append([[0, 0], truck_LW[i]])
                price_list.append(truck_price[i])
                count -= 1

        vol_total += x.varValue * truck_vol[i] / (1 - buffer_func)

    vol_total = round(vol_total, 2)
    return (
     truck_list, 1 - vol / vol_total, price_list)


def LP_layer_truck(Pallet=[], truck=[], truck_h_temp=[], floor=[], limited_H=[], N=[], H=[], rack_LW=[], rack_H=[], bf_c=0, bf_f=1, step_sol_temp=0):
    Combo_Pallet = []
    Pallet_Num = []
    Pile_Num = []
    rack_pile_num = []
    for i in range(len(truck)):
        Combo_Pallet.append(L_arrange(limited_H[i], H))

    if (Combo_Pallet == [[]]) & (rack_LW == [[]]):
        return (-1, [], [], [])
    else:
        if rack_LW != []:
            Combo_Rack = []
            for j in range(len(truck)):
                Combo_Rack.append([])
                for r in range(len(rack_LW)):
                    if isinstance(rack_H[r], list):
                        rack_temp = rack_H[r]
                    else:
                        rack_temp = [
                         rack_H[r]]
                    Combo_Rack[j].append(L_arrange(truck_h_temp[j], rack_temp))

            if Combo_Rack == [[]]:
                return (-1, [], [])
        else:
            prob1 = pp.LpProblem('Layer_truck', pp.LpMinimize)
            variables = []
            for j in range(len(truck)):
                for l in range(len(Pallet)):
                    count = len(Combo_Pallet[j])
                    if count < 11:
                        variables += [pp.LpVariable((str(j) + 'T' + str(l) + 'P' + 'C' + str(1) + '%d' % i), lowBound=0, cat=(pp.LpInteger)) for i in range(count)]
                    else:
                        while count > 0:
                            n = int((len(Combo_Pallet[j]) - count) / 10)
                            str_num = 1 * 10 ** n
                            for k in range(n):
                                str_num += 9 * 10 ** k

                            if count >= 10:
                                variables += [pp.LpVariable((str(j) + 'T' + str(l) + 'P' + 'C' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(10)]
                            else:
                                variables += [pp.LpVariable((str(j) + 'T' + str(l) + 'P' + 'C' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(count)]
                            count -= 10

            truck_rack = truck.copy()
            for j in range(len(truck_rack)):
                for r in range(len(rack_LW)):
                    count_rack = len(Combo_Rack[j][r])
                    if count_rack < 11:
                        variables += [pp.LpVariable((str(j) + 'T' + str(r) + 'R' + 'C' + str(1) + '%d' % i), lowBound=0, cat=(pp.LpInteger)) for i in range(count_rack)]
                    else:
                        while count_rack > 0:
                            n = int((len(Combo_Rack[j][r]) - count_rack) / 10)
                            str_num = 1 * 10 ** n
                            for k in range(n):
                                str_num += 9 * 10 ** k

                            if count_rack >= 10:
                                variables += [pp.LpVariable((str(j) + 'T' + str(r) + 'R' + 'C' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(10)]
                            else:
                                variables += [pp.LpVariable((str(j) + 'T' + str(r) + 'R' + 'C' + str(str_num) + '%d' % (i % 10)), lowBound=0, cat=(pp.LpInteger)) for i in range(count_rack)]
                            count_rack -= 10

            count_matrix = []
            count_matrix_cum = []
            count_num = 0
            for j in range(len(truck)):
                count_matrix.append([])
                for l in range(len(Pallet)):
                    count_matrix[j].append(len(Combo_Pallet[j]))
                    count_num += len(Combo_Pallet[j])
                    count_matrix_cum.append(count_num)

            count_matrix_cum.insert(0, 0)
            obj_temp = 0
            var_count = -1
            for j in range(len(truck)):
                for l in range(len(Pallet)):
                    for c in range(len(Combo_Pallet[j])):
                        var_count += 1
                        obj_temp += variables[var_count] * Pallet[l][0] * Pallet[l][1] / floor_list[j]

            count_pallet = var_count + 1
            for j in range(len(truck)):
                for r in range(len(rack_LW)):
                    for c in range(len(Combo_Rack[j][r])):
                        var_count += 1
                        obj_temp += variables[var_count] * rack_LW[r][0] * rack_LW[r][1]

            obj = obj_temp
            prob1 += obj
            count_step_rack = count_pallet
            for j in range(len(truck)):
                temp1 = 0
                for l in range(len(Pallet)):
                    temp1 += sum(variables[i] * (Pallet[l][0] * Pallet[l][1]) for i in range(count_matrix_cum[j * len(Pallet) + l], count_matrix_cum[j * len(Pallet) + l + 1]))

                temp2 = 0
                for r in range(len(rack_LW)):
                    temp2 += sum(variables[ii] * floor[j] * rack_LW[r][0] * rack_LW[r][1] for ii in range(count_step_rack, count_step_rack + len(Combo_Rack[j][r])))
                    count_step_rack = count_step_rack + len(Combo_Rack[j][r])

                temp = temp1 + temp2
                prob1 += temp - floor[j] * (truck[j][1][0] * truck[j][1][1] * (1 - bf_c)) <= 0
                prob1 += temp - floor[j] * (truck[j][1][0] * truck[j][1][1] * (1 - bf_f)) >= 0

            Combo_Pallet_Expand = []
            for j in range(len(truck)):
                Combo_Pallet_Expand.append([])
                for l in range(len(Pallet)):
                    Combo_Pallet_Expand[j].append(Combo_Pallet[j])

            temp_list_pal = []
            for l in range(len(Pallet)):
                temp_list_pal.append([])
                for h in range(len(H)):
                    temp_list_pal[l].append(0)

            count_var = -1
            for j in range(len(truck)):
                for l in range(len(Pallet)):
                    for c in range(len(Combo_Pallet_Expand[j][l])):
                        count_var += 1
                        for h in range(len(H)):
                            temp_list_pal[l][h] += Combo_Pallet_Expand[j][l][c].count(H[h]) * variables[count_var]

            for l in range(len(Pallet)):
                for h in range(len(H)):
                    prob1 += temp_list_pal[l][h] - N[l][h] >= 0

            temp_list_rack = []
            for r in range(len(rack_LW)):
                temp_list_rack.append([])
                for h in range(len(rack_H[r])):
                    temp_list_rack[r].append(0)

            for j in range(len(truck)):
                for r in range(len(rack_LW)):
                    for c in range(len(Combo_Rack[j][r])):
                        count_var += 1
                        for h in range(len(rack_H[r])):
                            temp_list_rack[r][h] += Combo_Rack[j][r][c].count(rack_H[r][h]) * variables[count_var]

            for r in range(len(rack_LW)):
                for h in range(len(rack_H[r])):
                    prob1 += temp_list_rack[r][h] - rack_Num[r][h] >= 0

            status = prob1.solve()
            if status == 1:
                print('')
                print('####################################### Layer truck ILP Detail Start ##################################')
                print(prob1)
                print('######################################## Layer truck ILP Detail End ###################################')
                print('')
                count_var = -1
                for j in range(len(truck)):
                    Pallet_Num.append([])
                    for l in range(len(Pallet)):
                        Pallet_Num[j].append(0)
                        for c in range(len(Combo_Pallet_Expand[j][l])):
                            count_var += 1
                            if variables[count_var].varValue != 0:
                                print('We need ', variables[count_var].varValue, ' \tpiles of mode: \tOn truck ', j + 1, '(', truck[j][1][0], 'm truck:)', '\t On Pallet ', Pallet[l], ',\t', Combo_Pallet_Expand[j][l][c])
                                height_in_mode = list(Combo_Pallet_Expand[j][l][c])
                                height_in_mode = del_blk(height_in_mode)
                                height_in_mode = del_rdd(height_in_mode)
                                Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]), variables[count_var].varValue)
                                Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]) + 1, str(Pallet[l][0]) + 'm X ' + str(Pallet[l][1]) + 'm')
                                for cH in range(len(height_in_mode)):
                                    if height_in_mode[cH] > 0:
                                        Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]) + 2, height_in_mode[cH])
                                        Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]) + 3, 'Pallet')
                                        Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]) + 4, Combo_Pallet_Expand[j][l][c].count(height_in_mode[cH]))
                                        Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]) + 5, 'Truck' + str(j + 1) + '(' + str(truck[j][1][0]) + 'm)')
                                        step_sol_temp += 1

                                Pallet_Num[j][l] += variables[count_var].varValue

                print('')
                for j in range(len(truck)):
                    rack_pile_num.append([])
                    for r in range(len(rack_LW)):
                        rack_pile_num[j].append(0)
                        for c in range(len(Combo_Rack[j][r])):
                            count_var += 1
                            if variables[count_var].varValue != 0:
                                print('We need ', variables[count_var].varValue, ' \tpiles of mode: \tOn truck ', j + 1, '(', truck[j][1][0], 'm truck:)', '\t Of Rack   ', rack_LW[r], ',\t', Combo_Rack[j][r][c])
                                height_in_mode = list(Combo_Rack[j][r][c])
                                height_in_mode = del_blk(height_in_mode)
                                height_in_mode = del_rdd(height_in_mode)
                                Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]), variables[count_var].varValue)
                                Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]) + 1, str(rack_LW[r][0]) + 'm X ' + str(rack_LW[r][1]) + 'm')
                                for cH in range(len(height_in_mode)):
                                    if height_in_mode[cH] > 0:
                                        Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]) + 2, height_in_mode[cH])
                                        Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]) + 3, 'Rack')
                                        Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]) + 4, Combo_Rack[j][r][c].count(height_in_mode[cH]))
                                        Solution_Sheet.write(step_sol_temp + int(solution_row[0]), int(solution_col[0]) + 5, 'Truck' + str(j + 1) + '(' + str(truck[j][1][0]) + 'm)')
                                        step_sol_temp += 1

                                rack_pile_num[j][r] += variables[count_var].varValue

            if Pallet_Num != []:
                for j in range(len(truck)):
                    Pile_Num.append([])
                    for l in range(len(Pallet)):
                        Pile_Num[j].append(0)
                        Pile_Num[j][l] = Pallet_Num[j][l] / floor[j]

        return (
         prob1, Pallet_Num, Pile_Num, rack_pile_num)


def Distri_item(type_list, contain_list, item_list):
    item_area = []
    item_total_area = 0
    item_allocate_area_max = []
    item_allocate_area = []
    item_allocate_list = []
    area = []
    total_area = 0
    allocate_ratio = []
    temp_list = []
    for i, it in enumerate(item_list):
        item_area.append(it[0] * it[1])
        item_total_area += it[0] * it[1]

    for i, cl in enumerate(contain_list):
        area.append(cl[1][0] * cl[1][1])
        total_area += cl[1][0] * cl[1][1]

    for i, cl in enumerate(contain_list):
        allocate_ratio.append(round(area[i] / total_area, 4))

    for i in range(len(contain_list)):
        item_allocate_area_max.append(item_total_area * allocate_ratio[i])
        item_allocate_area.append(0)
        item_allocate_list.append([])

    while item_list != []:
        for i, pal in enumerate(item_allocate_list):
            if item_allocate_area[i] < item_allocate_area_max[i]:
                item_temp = item_list.pop(0)
                item_allocate_list[i].append(item_temp)
                item_allocate_area[i] += item_area.pop(0)
                if item_list == []:
                    break

    return item_allocate_list


import time

def now():
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))


print('TLP Model Launching....')
print('PC&L Costing Calculating & Solution Generating...')
workbook_temp = xlrd.open_workbook('Temp.xls', formatting_info=True)
workbook_ui = xlrd.open_workbook('TLP_Solution_UI.xlsx')
Sheet_op = workbook_ui.sheet_by_name('Output')
Sheet_rc = workbook_ui.sheet_by_name('Rack Calculation')
Sheet_cp = workbook_ui.sheet_by_name('Container Panel')
Sheet_strategy = workbook_ui.sheet_by_name('Strategy')
Sheet_tp_op = workbook_temp.sheet_by_name('Output')
Sheet_tp_sl = workbook_temp.sheet_by_name('solution')
workbook_temp_copy = copy.copy(workbook_temp)
Temp_Sheet = workbook_temp_copy.get_sheet(0)
Solution_Sheet = workbook_temp_copy.get_sheet(1)
timeout_set = Sheet_strategy.col_values(9, 1, 2)[0]
N2_fill_YoN = Sheet_strategy.col_values(1, 2, 3)[0]
clear_content_count = Sheet_tp_op.col_values(13, 0, 2)
for i in range(int(clear_content_count[0]), int(clear_content_count[1])):
    for j in range(10):
        Temp_Sheet.write(i, j, '')

solution_row = Sheet_tp_sl.row_values(1, 33, 35)
solution_col = Sheet_tp_sl.row_values(1, 36, 38)
for i in range(int(solution_row[0]), int(solution_row[1])):
    for j in range(int(solution_col[0]), int(solution_col[1])):
        Solution_Sheet.write(i, j, '')

box_type_byparts = Sheet_op.col_values(2, 2, 61)
box_num_byparts = Sheet_op.col_values(3, 2, 61)
box_category_bybox = Sheet_op.col_values(4, 2, 61)
box_price_bybox = Sheet_op.col_values(5, 2, 61)
box_h_bybox = Sheet_op.col_values(8, 2, 61)
Box_H = del_rdd(box_h_bybox)
Box_H = del_blk(Box_H)
Box_H = list(reversed(sorted(Box_H)))
H_Complete = []
for i in range(len(Box_H)):
    H_Complete.append(1)

Num_RB_AN1_bybox = Sheet_op.col_values(9, 2, 61)
L_RB_AN1_bybox = Sheet_op.col_values(6, 2, 61)
W_RB_AN1_bybox = Sheet_op.col_values(7, 2, 61)
Num_RB_AN1_bybox, box_type_byparts, box_price_bybox, L_RB_AN1_bybox, W_RB_AN1_bybox, box_h_bybox, box_num_byparts, box_category_bybox = del_blank_queue(Num_RB_AN1_bybox, box_type_byparts, box_price_bybox, L_RB_AN1_bybox, W_RB_AN1_bybox, box_h_bybox, box_num_byparts, box_category_bybox, length=8)
box_type_temp = box_type_byparts.copy()
box_type_temp, Box_Price_ByType = type_match_unique(box_type_temp, box_price_bybox)
box_type_temp = box_type_byparts.copy()
box_type_temp, Box_Cat_ByType = type_match_unique(box_type_temp, box_category_bybox)
box_type_temp = box_type_byparts.copy()
box_type_temp, L_RB_AN1_ByType = type_match_unique(box_type_temp, L_RB_AN1_bybox)
box_type_temp = box_type_byparts.copy()
box_type_temp, W_RB_AN1_ByType = type_match_unique(box_type_temp, W_RB_AN1_bybox)
box_type_temp = box_type_byparts.copy()
box_type_temp, H_RB_AN1_ByType = type_match_unique(box_type_temp, box_h_bybox)
box_type_temp = box_type_byparts.copy()
box_type_temp, Num_RB_AN1_ByType = type_match_unique(box_type_temp, Num_RB_AN1_bybox)
box_type_temp = box_type_byparts.copy()
box_type_temp, Box_Num_ByType = type_add(box_type_temp, box_num_byparts)
Box_Type = box_type_temp
for i in range(len(Box_Type)):
    Temp_Sheet.write(22 + i, 0, Box_Type[i])
    Temp_Sheet.write(22 + i, 1, Box_Num_ByType[i])
    Temp_Sheet.write(22 + i, 2, Box_Price_ByType[i])
    Temp_Sheet.write(22 + i, 4, Box_Cat_ByType[i])

Num_of_Pallet_Type = Sheet_op.row_values(2, 0, 1)
Num_of_Pallet_Type = int(Num_of_Pallet_Type[0])
num_truck_std = Sheet_cp.row_values(0, 17, 18)
num_pal_std = Sheet_cp.row_values(0, 19, 20)
num_box_std = Sheet_cp.row_values(0, 21, 22)
num_truck_adh = Sheet_cp.row_values(1, 17, 18)
num_pal_adh = Sheet_cp.row_values(1, 19, 20)
num_box_adh = Sheet_cp.row_values(1, 21, 22)
num_truck_std = int(num_truck_std[0])
num_pal_std = int(num_pal_std[0])
num_box_std = int(num_box_std[0])
num_truck_adh = int(num_truck_adh[0])
num_pal_adh = int(num_pal_adh[0])
num_box_adh = int(num_box_adh[0])
loc_truck_start = Sheet_cp.row_values(3, 17, 18)
loc_pal_std = Sheet_cp.row_values(3, 19, 20)
loc_box_std = Sheet_cp.row_values(3, 21, 22)
loc_truck_end = Sheet_cp.row_values(4, 17, 18)
loc_pal_adh = Sheet_cp.row_values(4, 19, 20)
loc_box_adh = Sheet_cp.row_values(4, 21, 22)
loc_truck_start = int(loc_truck_start[0])
loc_pal_std = int(loc_pal_std[0])
loc_box_std = int(loc_box_std[0])
loc_truck_end = int(loc_truck_end[0])
loc_pal_adh = int(loc_pal_adh[0])
loc_box_adh = int(loc_box_adh[0])
box_type_avaliable = Sheet_op.col_values(11, 2, 61)
palletlayer_h_bybox = Sheet_op.col_values(12, 2, 61)
palletlayer_type_bybox_N1 = Sheet_op.col_values(15, 2, 61)
palletlayer_num_bybox_N1 = Sheet_op.col_values(13, 2, 61)
BoxInPallet_Num_bybox_N1 = Sheet_op.col_values(19, 2, 61)
Pallet_Type_avail = []
Pallet_LW_avail = []
Pallet_Price_avail = []
for i in range(num_pal_std):
    Pallet_Type_avail.append(Sheet_cp.row_values(loc_pal_std + i, 1, 2))
    Pallet_LW_avail.append(Sheet_cp.row_values(loc_pal_std + i, 5, 7))
    Pallet_Price_avail.append(Sheet_cp.row_values(loc_pal_std + i, 8, 9))

for i in range(num_pal_adh):
    Pallet_Type_avail.append(Sheet_cp.row_values(loc_pal_adh + i, 1, 2))
    Pallet_LW_avail.append(Sheet_cp.row_values(loc_pal_adh + i, 5, 7))
    Pallet_Price_avail.append(Sheet_cp.row_values(loc_pal_adh + i, 8, 9))

palletlayer_type_bybox_N1_temp = palletlayer_type_bybox_N1.copy()
palletlayer_h_bybox_temp = palletlayer_h_bybox.copy()
Palletlayer_Type_N1, PalletLayer_H_groupby_Pallet_N1 = type_match_group(palletlayer_type_bybox_N1_temp, palletlayer_h_bybox_temp)
palletlayer_type_bybox_N1_temp = palletlayer_type_bybox_N1.copy()
palletlayer_h_bybox_temp = palletlayer_h_bybox.copy()
Palletlayer_Type_N1, PalletLayer_H_groupby_Pallet_N1_repeat = type_match_group(palletlayer_type_bybox_N1_temp, palletlayer_h_bybox_temp, 0)
palletlayer_type_bybox_N1_temp = palletlayer_type_bybox_N1.copy()
box_type_avaliable_temp = box_type_avaliable.copy()
Palletlayer_Type_N1, Box_Type_groupby_Pallet_N1 = type_match_group(palletlayer_type_bybox_N1_temp, box_type_avaliable_temp)
palletlayer_type_bybox_N1_temp = palletlayer_type_bybox_N1.copy()
palletlayer_num_bybox_N1_temp = palletlayer_num_bybox_N1.copy()
Palletlayer_Type_N1, Box_Laynum_groupby_Pallet_N1 = type_match_group(palletlayer_type_bybox_N1_temp, palletlayer_num_bybox_N1_temp, 0)
palletlayer_type_bybox_N1_temp = palletlayer_type_bybox_N1.copy()
BoxInPallet_Num_bybox_N1_temp = BoxInPallet_Num_bybox_N1.copy()
Palletlayer_Type_N1, Box_inLayer_groupby_Pallet_N1 = type_match_group(palletlayer_type_bybox_N1_temp, BoxInPallet_Num_bybox_N1_temp, 0)
Palletlayer_LW_N1 = []
for i in range(len(Palletlayer_Type_N1)):
    for j in range(len(Pallet_Type_avail)):
        if Palletlayer_Type_N1[i] == Pallet_Type_avail[j][0]:
            Palletlayer_LW_N1.append(Pallet_LW_avail[j])

step_sol = 0
H_pallet = list(reversed(sorted(PalletLayer_H_groupby_Pallet_N1)))
for i in range(len(Palletlayer_Type_N1)):
    for k in range(len(H_pallet[i])):
        for j in range(len(Box_Type_groupby_Pallet_N1[i])):
            if PalletLayer_H_groupby_Pallet_N1_repeat[i][j] == H_pallet[i][k] and int(Box_Laynum_groupby_Pallet_N1[i][j]) > 0:
                Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]), int(Box_Laynum_groupby_Pallet_N1[i][j]))
                Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]) + 1, str(Palletlayer_LW_N1[i][0]) + 'm X ' + str(Palletlayer_LW_N1[i][1]) + 'm')
                Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]) + 2, PalletLayer_H_groupby_Pallet_N1_repeat[i][j])
                Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]) + 3, str(Box_Type_groupby_Pallet_N1[i][j]))
                Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]) + 4, int(Box_inLayer_groupby_Pallet_N1[i][j]))
                Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]) + 5, 'n1')
                step_sol += 1

n1 = []
n2 = []
for j in range(len(Pallet_Type_avail)):
    n1.append([])
    for k in range(len(Box_H)):
        n1[j].append(0)
        for i in range(len(palletlayer_type_bybox_N1)):
            if [
             palletlayer_type_bybox_N1[i]] == Pallet_Type_avail[j] and round(palletlayer_h_bybox[i], 5) == round(Box_H[k], 5):
                n1[j][k] += round(palletlayer_num_bybox_N1[i])

Container_Type = Sheet_rc.col_values(10, 3, 62)
Rack_Type = Sheet_rc.col_values(3, 3, 62)
Rack_Num = Sheet_rc.col_values(4, 3, 62)
Rack_L = Sheet_rc.col_values(5, 3, 62)
Rack_W = Sheet_rc.col_values(6, 3, 62)
Rack_H = Sheet_rc.col_values(7, 3, 62)
Rack_Price = Sheet_rc.col_values(8, 3, 62)
Container_Type, Rack_Type, Rack_Num, Rack_L, Rack_W, Rack_H, Rack_Price = del_blank_queue(Container_Type, Rack_Type, Rack_Num, Rack_L, Rack_W, Rack_H, Rack_Price, length=7)
row_count = 0
for i in range(len(Rack_Type)):
    if Rack_Num[i] > 0:
        Temp_Sheet.write(22 + len(Box_Type) + row_count, 0, Rack_Type[i])
        Temp_Sheet.write(22 + len(Box_Type) + row_count, 1, Rack_Num[i])
        Temp_Sheet.write(22 + len(Box_Type) + row_count, 2, Rack_Price[i])
        Temp_Sheet.write(22 + len(Box_Type) + row_count, 4, 'Rack')
        row_count += 1

Truck_Type_avail = Sheet_cp.col_values(1, loc_truck_start, loc_truck_end)
Truck_L = Sheet_cp.col_values(5, loc_truck_start, loc_truck_end)
Truck_W = Sheet_cp.col_values(6, loc_truck_start, loc_truck_end)
Truck_H = Sheet_cp.col_values(7, loc_truck_start, loc_truck_end)
Truck_Price = Sheet_cp.col_values(8, loc_truck_start, loc_truck_end)
Truck_LimH = Sheet_cp.col_values(9, loc_truck_start, loc_truck_end)
Truck_Floor = Sheet_cp.col_values(10, loc_truck_start, loc_truck_end)
Truck_LW = []
Truck_Area = []
LW_RB_groupby_H = []
Num_RB_groupby_H = []
Type_RB_groupby_H = []
Truck_Type_avail, Truck_L, Truck_W, Truck_H, Truck_Price, Truck_LimH, Truck_Floor = del_blank_queue(Truck_Type_avail, Truck_L, Truck_W, Truck_H, Truck_Price, Truck_LimH, Truck_Floor, length=7)
Truck_LW = OneD_To_TwoD(Truck_L, Truck_W, Truck_LW)
print('area info matching...')
print('price info matching...')
for i, tr in enumerate(Truck_LW):
    Truck_Area.append(round(tr[0] * tr[1], 2))

time.sleep(0.5)
print('matching completed!')
print('remaining box categorying...( by height,2D info and num recorded )')
time.sleep(0.5)
for i in range(len(Box_H)):
    LW_temp = []
    num_temp = []
    type_temp = []
    for j in range(len(Num_RB_AN1_ByType)):
        if H_RB_AN1_ByType[j] == Box_H[i]:
            LW_temp.append([L_RB_AN1_ByType[j], W_RB_AN1_ByType[j]])
            num_temp.append(round(Num_RB_AN1_ByType[j]))
            type_temp.append(Box_Type[j])

    LW_RB_groupby_H.append(LW_temp)
    Num_RB_groupby_H.append(num_temp)
    Type_RB_groupby_H.append(type_temp)

del LW_temp
del i
del j
del num_temp
del L_RB_AN1_ByType
del H_RB_AN1_ByType
del W_RB_AN1_ByType
print('categorying completed ')
Print_CMD('N1', 'the layers arranged by same type of box', n1)
Combo_N2_byP_byH = []
Num_Combo_N2_byP_byH = []
for p in range(Num_of_Pallet_Type):
    pallet_act = Pallet_LW_avail[p]
    p_layer_add_n2_temp = []
    Combo_N2_byP_byH.append([])
    Num_Combo_N2_byP_byH.append([])
    for h in range(len(LW_RB_groupby_H)):
        combo_inH = []
        LW_boxlist_from_combo_inH = []
        LW_boxlist_inH = LW_RB_groupby_H[h]
        num_boxlist_inH = Num_RB_groupby_H[h]
        type_boxlist_inH = Type_RB_groupby_H[h]
        Combo_N2_byP_byH[p].append([])
        Num_Combo_N2_byP_byH[p].append([])
        if N2_fill_YoN == 0:
            N2_mode = 0
            box_typenum = 1
        else:
            N2_mode = 1
            tolerance_1 = 0
            box_typenum = 0
        if len(LW_boxlist_inH) > box_typenum:
            combo_inH = P_arrange(pallet_act, LW_boxlist_inH, N2_mode)
            for i1 in range(len(combo_inH)):
                combo_inH[i1] = list(combo_inH[i1])

            for i2 in range(len(num_boxlist_inH)):
                num_boxlist_inH[i2] = int(num_boxlist_inH[i2])

            combo_inH = Combo_Filter(combo_inH, num_boxlist_inH, tolerance_1)
            LW_boxlist_from_combo_inH = Box_List(combo_inH, LW_boxlist_inH, [])
            for k in range(len(LW_boxlist_from_combo_inH)):
                temp_boxlist = LW_boxlist_from_combo_inH[k]
                the_realm = [[0, 0], pallet_act]
                items_sorted = list(reversed(sorted(temp_boxlist)))
                bestfit, boxremain = Box_Master(items_sorted, the_realm)
                if boxremain != []:
                    combo_inH[k] = -1

            combo_inH = list(filter((lambda x: x != -1), combo_inH))
            Combo_N2_byP_byH[p][h] += combo_inH.copy()

Combo_N2_ByH = []
for i in range(len(LW_RB_groupby_H)):
    Combo_N2_ByH.append([])
    for j in range(Num_of_Pallet_Type):
        Combo_N2_ByH[i].append(Combo_N2_byP_byH[j][i])

temp_N2 = []
for i in range(len(Combo_N2_ByH)):
    Num_RB_inH_N2 = Num_RB_groupby_H[i]
    Type_RB_inH_N2 = Type_RB_groupby_H[i]
    Combo_N2_inH_byP = Combo_N2_ByH[i]
    Combo_N2_inH_1D, Position_in_1D_inH_byP = Combo_2D_to_1D(Combo_N2_inH_byP)
    num_boxlist_errproof = remove_0([], Num_RB_inH_N2, Combo_N2_inH_1D)
    if N2_fill_YoN == 0:
        var_combo_N2_1D, Num_RB_AN2, prob_n2 = ILP_Box(Combo_N2_inH_1D, Num_RB_inH_N2, num_boxlist_errproof)
    else:
        var_combo_N2_1D, Num_RB_AN2, prob_n2 = ILP_Box_nonleft(Combo_N2_inH_1D, Num_RB_inH_N2, num_boxlist_errproof)
    for b in range(len(Num_RB_groupby_H[i])):
        if Combo_N2_inH_1D != []:
            Num_RB_groupby_H[i][b] -= round(sum(Combo_N2_inH_1D[con][b] * var_combo_N2_1D[con] for con in range(len(var_combo_N2_1D))))

    if None not in var_combo_N2_1D:
        if sum(var_combo_N2_1D[n] for n in range(len(var_combo_N2_1D))) > 0:
            var_combo_N2_2D = List_1D_to_2D(var_combo_N2_1D, Combo_N2_ByH[i])
            for p in range(len(Pallet_LW_avail)):
                for c in range(len(var_combo_N2_2D[p])):
                    if var_combo_N2_2D[p][c] > 0:
                        Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]), int(var_combo_N2_2D[p][c]))
                        Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]) + 1, str(Pallet_LW_avail[p][0]) + 'm X ' + str(Pallet_LW_avail[p][1]) + 'm')
                        Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]) + 2, Box_H[i])
                        for b in range(len(Type_RB_groupby_H[i])):
                            if Combo_N2_ByH[i][p][c][b] > 0:
                                Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]) + 3, str(Type_RB_groupby_H[i][b]))
                                Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]) + 4, int(Combo_N2_ByH[i][p][c][b]))
                                Solution_Sheet.write(step_sol + int(solution_row[0]), int(solution_col[0]) + 5, 'n2')
                                step_sol += 1

        step_add = 0
        P_layer_add_byP_inH_temp = []
        for j in range(len(Position_in_1D_inH_byP)):
            step_add_byP = 0
            P_layer_add_byP_inH_temp.append(0)
            while step_add_byP < Position_in_1D_inH_byP[j]:
                P_layer_add_byP_inH_temp[j] += var_combo_N2_1D[step_add]
                step_add_byP += 1
                step_add += 1

        temp_N2.append(P_layer_add_byP_inH_temp)

for i in range(Num_of_Pallet_Type):
    n2.append([])
    for j in range(len(LW_RB_groupby_H)):
        n2[i].append([])

for i in range(len(temp_N2)):
    for j in range(len(temp_N2[i])):
        n2[j][i] = temp_N2[i][j]

Print_CMD('N2', 'the layers arranged by the remaining box after N1 with same height', n2)
LW_RB_N3 = []
Num_RB_N3 = []
for i in range(len(LW_RB_groupby_H)):
    for j in range(len(LW_RB_groupby_H[i])):
        if LW_RB_groupby_H[i][j] not in LW_RB_N3:
            LW_RB_N3.append(LW_RB_groupby_H[i][j])
            if Num_RB_groupby_H[i][j] >= 0:
                Num_RB_N3.append(round(Num_RB_groupby_H[i][j]))
            else:
                Num_RB_N3.append(0)
        else:
            for k in range(len(LW_RB_N3)):
                if LW_RB_groupby_H[i][j] == LW_RB_N3[k] and Num_RB_groupby_H[i][j] > 0:
                    Num_RB_N3[k] += round(Num_RB_groupby_H[i][j])

Combo_N3_byP = []
for p in range(Num_of_Pallet_Type):
    Combo_N3_byP.append([])
    pallet_act = Pallet_LW_avail[p]
    combo_temp_inP = P_arrange(pallet_act, LW_RB_N3, 1)
    combo_temp_inP = Combo_Filter(combo_temp_inP, Num_RB_N3)
    lw_boxlist_temp_inP = Box_List(combo_temp_inP, LW_RB_N3, [])
    for k in range(len(lw_boxlist_temp_inP)):
        temp_boxlist = lw_boxlist_temp_inP[k]
        the_realm = [[0, 0], pallet_act]
        items_sorted = list(reversed(sorted(temp_boxlist)))
        bestfit, boxremain = Box_Master(items_sorted, the_realm)
        if boxremain != []:
            combo_temp_inP[k] = -1

    combo_temp_inP = list(filter((lambda x: x != -1), combo_temp_inP))
    Combo_N3_byP[p] += combo_temp_inP

Combo_N3_1D, Position_in_1D_byP = Combo_2D_to_1D(Combo_N3_byP)
num_boxlist_errproof = remove_0([], Num_RB_N3, Combo_N3_1D)
var_combo_N3_1D, Num_RB_AN3, prob_n3 = ILP_Box_nonleft(Combo_N3_1D, Num_RB_N3, num_boxlist_errproof)
n3 = []
n3, RB_AN3_LW_H = N3_Generator(var_combo_N3_1D, Position_in_1D_byP, Combo_N3_1D, LW_RB_groupby_H, Num_RB_groupby_H, LW_RB_N3, Box_H)
Print_CMD('N3', 'the layers arranged by all the remaining box after N1&N2', n3)
print('all box arranged into pallets')
print('')
print('Linear Programming Launched ### if it has taken too much time,errors may exist in excel input,please check ###')
for i in range(len(n1)):
    n1[i] = n1[i][0:len(Box_H)]
    for j in range(len(n1[i])):
        n1[i][j] = int(n1[i][j])

N = np.array(n1) + np.array(n2) + np.array(n3)
Pallet_Type_needed = Del_Pallet(N, Pallet_Type_avail)
Pallet_LW_needed = type_match_sub(Pallet_Type_needed, Pallet_Type_avail, Pallet_LW_avail)
Pallet_Price_needed = type_match_sub(Pallet_Type_needed, Pallet_Type_avail, Pallet_Price_avail)
N_remove0 = type_match_sub(Pallet_Type_needed, Pallet_Type_avail, N)
minus = len(N) - len(N_remove0)
Num_of_Pallet_Type_needed = Num_of_Pallet_Type - minus
floor = Truck_Floor
limited_H = Truck_LimH
truck_h = Truck_H
buffer = Sheet_strategy.col_values(5, 1, 2)[0]
buffer_2D_max = Sheet_strategy.col_values(1, 1, 2)[0]
buffer_2D_min = Sheet_strategy.col_values(3, 1, 2)[0]
visual_aid = Sheet_strategy.col_values(5, 2, 3)[0]
complete = 0
buffer_step = Sheet_strategy.col_values(7, 1, 2)[0]
success = []
step = 1
buffer_ceil = buffer_2D_max
buffer_floor = buffer_2D_min
buffer_interval = buffer_floor - buffer_ceil
renew = 0
while (complete < 1) | (success == []):
    for i in range(step_sol + int(solution_row[0]), int(solution_row[1])):
        for j in range(int(solution_col[0]), int(solution_col[1])):
            Solution_Sheet.write(i, j, '')

    if renew == 1:
        buffer = buffer_copy
        buffer *= buffer_step
        buffer_ceil = buffer_2D_max
        buffer_floor = buffer_2D_min
    vol1_Pallet = 0
    for l in range(len(Pallet_LW_needed)):
        for h in range(len(Box_H)):
            vol1_Pallet += N_remove0[l][h] * Box_H[h] * Pallet_LW_needed[l][0] * Pallet_LW_needed[l][1]

    vol1_Rack = 0
    for r in range(len(Rack_Type)):
        vol1_Rack += Rack_L[r] * Rack_W[r] * Rack_H[r] * Rack_Num[r]

    vol1 = vol1_Pallet + vol1_Rack
    truck_vol_real = []
    truck_vol_raw = []
    for j in range(len(Truck_LW)):
        truck_vol_real.append(Truck_LW[j][0] * Truck_LW[j][1] * limited_H[j] * floor[j])
        truck_vol_raw.append(Truck_LW[j][0] * Truck_LW[j][1] * Truck_H[j])

    Truck_LW1 = Truck_LW
    truck_list = []
    floor_list = []
    limited_H_list = []
    truck_h_list = []
    truck_vol1 = truck_vol_real.copy()
    truck_list, ratio, price_list = LP_Truck(vol1, truck_vol1, Truck_Price, Truck_LW1, truck_list, buffer, price_list=[])
    for i in range(len(truck_list)):
        floor_list.append(0)
        limited_H_list.append(0)
        truck_h_list.append(0)

    for i in range(len(truck_list)):
        for j in range(len(Truck_LW)):
            if truck_list[i][1] == Truck_LW[j]:
                floor_list[i] = floor[j]
                limited_H_list[i] = limited_H[j]
                truck_h_list[i] = truck_h[j]

    rack_LW = []
    rack_H = []
    rack_Num = []
    for i in range(len(Rack_Type)):
        temp = [
         Rack_L[i], Rack_W[i]]
        temp = list(reversed(sorted(temp)))
        if temp not in rack_LW:
            rack_LW.append(temp)
            rack_H.append([])
            rack_Num.append([])

    for i in range(len(Rack_Type)):
        for j in range(len(rack_LW)):
            temp = [
             Rack_L[i], Rack_W[i]]
            temp = list(reversed(sorted(temp)))
            if rack_LW[j] == temp:
                rack_H[j].append(Rack_H[i])

    for i in range(len(Rack_Type)):
        for j in range(len(rack_LW)):
            temp = [
             Rack_L[i], Rack_W[i]]
            temp = list(reversed(sorted(temp)))
            if rack_LW[j] == temp:
                rack_Num[j].append(Rack_Num[i])

    step_sol_pile = step_sol + 1
    Solution_Sheet.write(step_sol_pile + int(solution_row[0]), int(solution_col[0]), 'Num')
    Solution_Sheet.write(step_sol_pile + int(solution_row[0]), int(solution_col[0]) + 1, 'Pile Type')
    Solution_Sheet.write(step_sol_pile + int(solution_row[0]), int(solution_col[0]) + 2, 'Pallet Height')
    Solution_Sheet.write(step_sol_pile + int(solution_row[0]), int(solution_col[0]) + 3, 'Category')
    Solution_Sheet.write(step_sol_pile + int(solution_row[0]), int(solution_col[0]) + 4, 'Arrangement')
    Solution_Sheet.write(step_sol_pile + int(solution_row[0]), int(solution_col[0]) + 5, 'Truck')
    step_sol_pile += 1
    prob, Pallet_Num, Pile_Num, Rack_Pile_Num = LP_layer_truck(Pallet_LW_needed, truck_list, truck_h_list, floor_list, limited_H_list, N_remove0, Box_H, rack_LW, rack_H, buffer_ceil, buffer_floor, step_sol_pile)
    if prob == -1:
        complete = prob
    else:
        complete = prob.status
    buffer_copy = ratio
    renew = 0
    if complete != 1:
        if round(buffer_ceil - buffer_interval, 1) >= 0:
            buffer_floor = round(buffer_ceil, 1)
            buffer_ceil -= buffer_interval
        else:
            if round(buffer_ceil, 1) > 0:
                buffer_floor = round(buffer_ceil, 1)
                buffer_ceil = 0
            else:
                renew = 1
        continue
    else:
        print('strategy used:')
        print('buffer floor:', round(buffer_floor, 1))
        print('buffer ceil:', round(buffer_ceil, 1))
        print('')
        print('the least space we need in trucks are: ', round(vol1, 2))
        print('')
        print('truck avaliable:')
        for i in range(len(Truck_LW)):
            print('\t', Truck_LW[i][0], '\tm truck', '\tdimension:', Truck_LW[i], '\tlimited height:', limited_H[i], '\tfloor:', floor[i], '\ttruck vol:', round(truck_vol_real[i]), '\tprice:', Truck_Price[i])

        print('')
        truckvol_list = []
        Vol_total = 0
        Vol_total_raw = 0
        print('we need', len(truck_list), 'truck:')
        for i in range(len(truck_list)):
            print('\t', truck_list[i][1][0], '\tm truck')
            for j in range(len(Truck_LW)):
                if truck_list[i][1] == Truck_LW[j]:
                    truckvol_list.append(truck_vol_real[j])
                    Vol_total += truck_vol_real[j]
                    Vol_total_raw += truck_vol_raw[j]

    LimH_Type, Truck_byLimH = type_match_group(limited_H_list, truck_list, 0)
    LimH_Type, PileNum_byLimH = type_add(limited_H_list, Pile_Num)
    LimH_Type, Rack_Pile_Num_byLimH = type_add(limited_H_list, Rack_Pile_Num)
    for i in range(len(LimH_Type)):
        for l in range(len(Pallet_LW_needed)):
            PileNum_byLimH[i][l] = math.ceil(PileNum_byLimH[i][l])

    Rack_LW = rack_LW
    Pallet_List_1 = Box_List(PileNum_byLimH, Pallet_LW_needed, [])
    Pallet_List_2 = Box_List(Rack_Pile_Num_byLimH, Rack_LW, [])
    Pallet_List = []
    for i in range(len(Pallet_List_1)):
        Pallet_List.append(Pallet_List_1[i] + Pallet_List_2[i])

    pallet_totallist_temp = []
    for i in range(len(Truck_byLimH)):
        if len(Truck_byLimH[i]) > 1:
            pallet_list_temp = Pallet_List[i].copy()
            pallet_list_temp = Distri_item(LimH_Type[i], Truck_byLimH[i], pallet_list_temp)
            Pallet_List[i] = pallet_list_temp
            pallet_totallist_temp += Pallet_List[i]
        else:
            pallet_totallist_temp.append(Pallet_List[i])

    Pallet_List = pallet_totallist_temp
    print('')
    for i in range(len(Pallet_List)):
        if i <= 8:
            str_i = str(0) + str(i + 1)
        else:
            str_i = str(i + 1)
        print('pallet list wait to load onto the truck ', str_i, '\t', truck_list[i][1][0], 'm truck:\t', Pallet_List[i])

    time.sleep(1)
    print('')
    print('pallet loading onto the truck...:')
    for i in range(len(truck_list)):
        items = Pallet_List[i]
        the_realm = truck_list[i]
        items = deci(items)
        the_realm = deci(the_realm)
        The_Judge = []
        items_sorted = list(reversed(sorted(items)))
        a, b = Box_Master(items_sorted, the_realm)
        print('pallet arranged in truck', i + 1, ':\t', a)
        print('pallet remain after truck', i + 1, ':\t', b)
        the_realm_L = the_realm[1][0]
        the_realm_W = the_realm[1][1]
        if visual_aid == 'Y':
            import matplotlib.pylab as plt
            fig = plt.figure(figsize=(10, 10))
            ax = fig.add_subplot(111, aspect='equal')
            ax.set_xticks([0, 1])
            if the_realm_L > the_realm_W:
                the_realm_W1 = the_realm_W / the_realm_L * 0.9
                the_realm_L1 = 0.9
            else:
                the_realm_L1 = the_realm_L / the_realm_W * 0.9
                the_realm_W1 = 0.9
            ax.add_patch(plt.Rectangle((0.05, 0.05), the_realm_L1, the_realm_W1, fill=False, color='red', linewidth=3))
            for i in a:
                rect = plt.Rectangle((i[0][0] / the_realm_L * 0.9 + 0.05, i[0][1] / the_realm_L * 0.9 + 0.05), ((i[1][0] - i[0][0]) / the_realm_L * 0.9), (0.9 * (i[1][1] - i[0][1]) / the_realm_L), edgecolor='black', facecolor='grey', linewidth=2)
                ax.add_patch(rect)

            plt.show()
        if b != []:
            success.append(0)
        else:
            success.append(1)
        if 0 in success:
            break

    if 0 not in success:
        print('trial', step, ' succeeded')
        print('truck capacity%: ', round(vol1 / Vol_total_raw * 100, 2), '%')
        Temp_Sheet.write(4, 9, str(round(vol1 / Vol_total_raw * 100, 2)) + '%')
    else:
        print('trial', step, ' failed')
        success = []
        step += 1
        renew = 1

Num_plate = []
if len(Pallet_Num) != 0:
    for i in range(len(Pallet_Num[0])):
        Num_plate.append(0)

for i in range(len(truck_list)):
    for j in range(len(Pallet_Num[i])):
        Num_plate[j] += Pallet_Num[i][j]

for p in range(Num_of_Pallet_Type_needed):
    Temp_Sheet.write(13 + p, 0, str(Pallet_LW_needed[p][0]) + 'm X ' + str(Pallet_LW_needed[p][1]) + 'm')
    Temp_Sheet.write(13 + p, 1, Num_plate[p])
    Temp_Sheet.write(13 + p, 2, Pallet_Price_needed[p][0])

truck_list_temp = truck_list.copy()
truck_type, truck_count = type_count(truck_list_temp, Truck_Type_avail, Truck_LW)
truck_list_temp = truck_list.copy()
price_list_temp = price_list.copy()
truck_list_temp, price_list_type = type_match_unique(truck_list_temp, price_list_temp)
for i in range(len(truck_type)):
    Temp_Sheet.write(4 + i, 0, truck_type[i])
    Temp_Sheet.write(4 + i, 1, truck_count[i])
    Temp_Sheet.write(4 + i, 2, price_list_type[i])

price = round(sum(price_list[i] for i in range(len(price_list))), 2)
print('total price for transportation:\t', price)
runon = 0
while runon == 0:
    try:
        workbook_temp_copy.save('Temp.xls')
        os.startfile('Temp.xls')
        runon = 1
    except:
        print('Please Close Temp File First,Then press Enter')
        runon = 0
        input()

input()